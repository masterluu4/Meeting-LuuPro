import streamlit as st
import streamlit.components.v1 as components
import requests
import json
import io
import re
import tempfile
import os
import time
from datetime import datetime
import google.generativeai as genai
from docxtpl import DocxTemplate
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ================= CẤU HÌNH TRANG =================
st.set_page_config(page_title="THACO Smart Meeting & AI Biên Bản", page_icon="🚗", layout="wide")

# Khởi tạo Session State
defaults = {
    "transcript_list": [],      # [{time, source, target, terms}]
    "raw_stream_input": "",     # Luồng văn bản đồng bộ từ Web Audio
    "glossary": {},             # Từ điển THACO dùng chung
    "tm": {},                   # Bộ nhớ dịch TM dùng chung
    "is_synced": False,         # Trạng thái đã tải dữ liệu lần đầu
    "available_models": [],     # Danh sách model lấy từ API Key
    "recap_data": {
        "tom_tat_chung": "",
        "noi_dung_thao_luan": "",
        "ket_luan_chi_dao": "",
        "action_items": [
            {"tt": "1", "cong_viec": "", "phu_trach": "", "thoi_han": "", "ghi_chu": ""}
        ]
    }
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ================= HÀM LÀM SẠCH URL & KẾT NỐI GOOGLE SHEETS =================
def clean_webapp_url(raw_url):
    """Loại bỏ hoàn toàn dấu ngoặc vuông [], ngoặc kép, khoảng trắng bị dính khi copy/paste"""
    if not raw_url:
        return ""
    s = str(raw_url).strip()
    match = re.search(r'https?://[^\s\[\]"\'<>()]+', s)
    if match:
        return match.group(0)
    return s.strip("[]\"'() ")

def sync_from_sheets(raw_url):
    webapp_url = clean_webapp_url(raw_url)
    if not webapp_url:
        return False, "⚠️ Chưa nhập Link Web App URL hợp lệ."
    try:
        res = requests.get(webapp_url, timeout=12)
        if res.status_code == 200:
            data = res.json()
            st.session_state.glossary = data.get("glossary", {})
            st.session_state.tm = data.get("tm", {})
            st.session_state.is_synced = True
            return True, f"✅ Đã đồng bộ: {len(st.session_state.glossary)} thuật ngữ & {len(st.session_state.tm)} câu TM."
    except Exception as e:
        return False, f"❌ Lỗi kết nối Google Sheets: {e}"
    return False, f"❌ Đường dẫn không phản hồi (HTTP {res.status_code})."

def save_full_glossary_to_sheets(raw_url, updated_glossary_dict):
    webapp_url = clean_webapp_url(raw_url)
    if not webapp_url:
        return False, "Chưa nhập Link Web App URL."
    try:
        res_get = requests.get(webapp_url, timeout=10)
        cloud_glossary = {}
        if res_get.status_code == 200:
            cloud_glossary = res_get.json().get("glossary", {})
        
        cloud_glossary.update(updated_glossary_dict)
        payload = {"action": "save_glossary", "data": cloud_glossary}
        res_post = requests.post(webapp_url, data=json.dumps(payload), timeout=12)
        if res_post.status_code == 200 and res_post.json().get("status") == "ok":
            st.session_state.glossary = cloud_glossary
            return True, f"✅ Đã lưu đồng bộ {len(cloud_glossary)} thuật ngữ lên Google Sheets!"
        return False, f"❌ Lỗi từ Apps Script: {res_post.text}"
    except Exception as e:
        return False, f"❌ Lỗi kết nối: {e}"

def upsert_tm_to_sheets(raw_url, new_tm_dict):
    webapp_url = clean_webapp_url(raw_url)
    if not webapp_url or not new_tm_dict:
        return False, "Không có dữ liệu hoặc thiếu URL."
    try:
        clean_dict = {str(k).strip(): str(v).strip() for k, v in new_tm_dict.items() if str(k).strip() and str(v).strip()}
        payload = {"action": "upsert_tm", "data": clean_dict}
        res = requests.post(webapp_url, data=json.dumps(payload), timeout=12)
        if res.status_code == 200 and res.json().get("status") == "ok":
            st.session_state.tm.update(clean_dict)
            return True, f"✅ Đã cập nhật {len(clean_dict)} câu vào Bộ nhớ đệm (TM) dùng chung!"
        return False, f"❌ Lỗi từ Apps Script: {res.text}"
    except Exception as e:
        return False, f"❌ Lỗi kết nối: {e}"

# ================= HÀM QUÉT MODEL KHẢ DỤNG & CHỌN MODEL TỐI ƯU =================
@st.cache_data(ttl=600, show_spinner=False)
def fetch_available_models(api_key):
    """Lấy danh sách các model thực tế khả dụng với API Key của người dùng"""
    if not api_key:
        return []
    try:
        genai.configure(api_key=api_key.strip())
        models = []
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods:
                name = m.name.replace("models/", "")
                models.append(name)
        return models
    except Exception:
        return []

def pick_optimal_model(models_list, task="recap"):
    """Tự động đề xuất model tối ưu: 'pro' cho Recap/Audio, 'flash' cho Dịch nhanh"""
    if not models_list:
        return "gemini-1.5-pro" if task == "recap" else "gemini-1.5-flash"
    
    if task == "recap":
        priorities = ["gemini-2.5-pro", "gemini-1.5-pro", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    else:
        priorities = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
        
    for p in priorities:
        for m in models_list:
            if m == p or m.startswith(p):
                return m
    return models_list[0]

# ================= HÀM DỊCH THUẬT AI + SO KHỚP GLOSSARY & TM =================
def translate_batch_with_thaco_memory(lines_with_time, mode, api_key, model_name, raw_url=""):
    results = []
    tm_to_upsert = {}

    for time_tag, clean_text in lines_with_time:
        if not clean_text:
            continue
        if mode == "Việt -> Việt (Chỉ ghi phụ đề)":
            results.append({"time": time_tag, "source": clean_text, "target": clean_text, "terms": []})
            continue

        if clean_text in st.session_state.tm:
            results.append({
                "time": time_tag,
                "source": clean_text,
                "target": st.session_state.tm[clean_text],
                "terms": ["⚡ TM Dùng Chung"]
            })
            continue

        matched = {}
        for vn_term, en_term in st.session_state.glossary.items():
            if vn_term.lower() in clean_text.lower() or en_term.lower() in clean_text.lower():
                matched[vn_term] = en_term

        if not api_key:
            results.append({
                "time": time_tag,
                "source": clean_text,
                "target": "[Chưa nhập Gemini API Key]",
                "terms": list(matched.keys())
            })
            continue

        direction = "tiếng Anh sang tiếng Việt" if mode == "Anh -> Việt" else "tiếng Việt sang tiếng Anh"
        glossary_instruction = ""
        if matched:
            glossary_instruction = f"BẮT BUỘC sử dụng đúng thuật ngữ chuyên ngành THACO sau: {json.dumps(matched, ensure_ascii=False)}. "

        try:
            genai.configure(api_key=api_key.strip())
            model = genai.GenerativeModel(model_name)
            prompt = (
                f"Bạn là chuyên gia phiên dịch kỹ thuật và quản lý chất lượng của THACO AUTO. "
                f"Hãy dịch câu hội thoại sau từ {direction} thật chuẩn xác, ngắn gọn, đúng văn phong kỹ thuật ô tô. "
                f"{glossary_instruction}"
                f"Chỉ trả về duy nhất câu bản dịch, không giải thích:\n{clean_text}"
            )
            res = model.generate_content(prompt)
            translated = res.text.strip()

            results.append({
                "time": time_tag,
                "source": clean_text,
                "target": translated,
                "terms": list(matched.keys())
            })

            if len(clean_text.split()) >= 3:
                tm_to_upsert[clean_text] = translated

        except Exception as e:
            results.append({"time": time_tag, "source": clean_text, "target": f"[Lỗi AI: {e}]", "terms": []})

    if raw_url and tm_to_upsert:
        upsert_tm_to_sheets(raw_url, tm_to_upsert)

    return results

# ================= HÀM TẠO FILE WORD MẪU CHUẨN TRÌNH KÝ =================
def generate_sample_docx():
    doc = Document()
    title = doc.add_heading("BIÊN BẢN CUỘC HỌP", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_info = doc.add_paragraph()
    p_info.add_run("1. Tên cuộc họp: ").bold = True
    p_info.add_run("{{ten_cuoc_hop}}\n")
    p_info.add_run("2. Thời gian: ").bold = True
    p_info.add_run("{{ngay_hop}}\n")
    p_info.add_run("3. Địa điểm / Hình thức: ").bold = True
    p_info.add_run("{{dia_diem}}\n")
    p_info.add_run("4. Chủ trì: ").bold = True
    p_info.add_run("{{chu_tri}}\n")
    p_info.add_run("5. Thành phần tham dự: ").bold = True
    p_info.add_run("{{thanh_phan}}")

    doc.add_heading("I. TÓM TẮT MỤC TIÊU & BỐI CẢNH", level=2)
    doc.add_paragraph("{{tom_tat_chung}}")

    doc.add_heading("II. NỘI DUNG THẢO LUẬN CHI TIẾT", level=2)
    doc.add_paragraph("{{noi_dung_thao_luan}}")

    doc.add_heading("III. KẾT LUẬN & CHỈ ĐẠO CỦA CHỦ TRÌ", level=2)
    doc.add_paragraph("{{ket_luan_chi_dao}}")

    doc.add_heading("IV. KẾ HOẠCH HÀNH ĐỘNG & PHÂN CÔNG", level=2)
    doc.add_paragraph("{{ke_hoach_hanh_dong}}")

    sign_table = doc.add_table(rows=2, cols=2)
    sign_table.cell(0, 0).paragraphs[0].add_run("NGƯỜI LẬP BIÊN BẢN").bold = True
    sign_table.cell(0, 1).paragraphs[0].add_run("CHỦ TRÌ CUỘC HỌP PHÊ DUYỆT").bold = True
    sign_table.cell(1, 0).paragraphs[0].add_run("\n\n\n{{nguoi_lap}}")
    sign_table.cell(1, 1).paragraphs[0].add_run("\n\n\n{{chu_tri}}")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# ================= THANH BÊN (SIDEBAR) =================
with st.sidebar:
    st.header("⚙️ Cấu Hình Hệ Thống")
    
    default_key = st.secrets.get("GEMINI_API_KEY", "") if "GEMINI_API_KEY" in st.secrets else ""
    default_url = st.secrets.get("WEBAPP_URL", "") if "WEBAPP_URL" in st.secrets else ""

    gemini_api_key = st.text_input("🔑 1. Gemini API Key:", value=default_key, type="password")
    
    # Quét danh sách Model khả dụng ngay khi có API Key
    if gemini_api_key:
        models_list = fetch_available_models(gemini_api_key)
        if models_list:
            st.success(f"✅ API Key hợp lệ ({len(models_list)} model khả dụng)")
            opt_recap = pick_optimal_model(models_list, task="recap")
            opt_trans = pick_optimal_model(models_list, task="translate")
            
            selected_recap_model = st.selectbox(
                "🧠 Model AI cho Recap & Nghe File Âm Thanh (Tối ưu):",
                options=models_list,
                index=models_list.index(opt_recap) if opt_recap in models_list else 0,
                help="Hệ thống đã tự động chọn model mạnh nhất (Pro/Flash) cho phân tích sâu và nghe file ghi âm."
            )
            selected_trans_model = st.selectbox(
                "⚡ Model AI cho Dịch Phụ Đề Nhanh:",
                options=models_list,
                index=models_list.index(opt_trans) if opt_trans in models_list else 0,
                help="Hệ thống tự động chọn dòng Flash để dịch nhanh theo thời gian thực."
            )
        else:
            st.error("❌ Không lấy được danh sách Model. Vui lòng kiểm tra lại API Key.")
            selected_recap_model = "gemini-1.5-pro"
            selected_trans_model = "gemini-1.5-flash"
    else:
        st.caption("Nhập API Key để hệ thống hiển thị các Model khả dụng.")
        selected_recap_model = "gemini-1.5-pro"
        selected_trans_model = "gemini-1.5-flash"

    st.divider()
    raw_webapp_Input = st.text_input("☁️ 2. Link Google Apps Script (Dùng chung):", value=default_url)
    webapp_url = clean_webapp_url(raw_webapp_Input)
    
    if webapp_url and not st.session_state.is_synced:
        ok, msg = sync_from_sheets(webapp_url)
        if ok:
            st.toast(msg)

    if st.button("🔄 Đồng bộ lại Glossary & TM từ Sheets", use_container_width=True, type="primary"):
        ok, msg = sync_from_sheets(webapp_url)
        if ok:
            st.success(msg)
        else:
            st.error(msg)

    st.caption(f"📊 Kho dữ liệu chung: **{len(st.session_state.glossary)}** thuật ngữ | **{len(st.session_state.tm)}** câu TM")

    st.divider()
    st.subheader("🌐 Chế độ Ngôn ngữ Họp")
    trans_mode = st.radio(
        "Chọn chiều nhận diện & dịch:",
        ["Anh -> Việt", "Việt -> Anh", "Việt -> Việt (Chỉ ghi phụ đề)"]
    )
    
    st.divider()
    with st.expander("➕ Thêm nhanh Thuật ngữ vào Kho chung"):
        new_vn = st.text_input("Từ gốc (Cột A - Tu_Goc):")
        new_en = st.text_input("Từ dịch (Cột B - Tu_Dich):")
        if st.button("💾 Lưu an toàn vào Glossary", use_container_width=True):
            if new_vn and new_en and webapp_url:
                ok, msg = save_full_glossary_to_sheets(webapp_url, {new_vn.strip(): new_en.strip()})
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
            else:
                st.warning("Vui lòng nhập đủ Từ gốc, Từ dịch và Link Web App URL.")

# ================= GIAO DIỆN CHÍNH =================
st.title("🚗 Hệ Thống Trợ Lý Họp, Dịch Thuật Real-time & Xuất Biên Bản THACO")

tab1, tab2, tab3, tab4 = st.tabs([
    "🎙️ Bước 1: Thu Âm VB-Cable & Phụ Đề",
    "🧠 Bước 2: AI Recap (Từ File Ghi Âm / .TXT)",
    "📄 Bước 3: Xuất Biên Bản Word Trình Ký",
    "📚 Quản Trị Kho TM / Glossary Chung"
])

# ----------------- TAB 1: THU ÂM VB-CABLE TRÊN WEB & PHỤ ĐỀ -----------------
with tab1:
    st.info(
        "🔊 **Thiết lập chuẩn với VB-Audio Virtual Cable:** "
        "Trong cài đặt âm thanh của **MS Teams / Zoom / VooV**, chọn Loa (Speaker) là `CABLE Input`. "
        "Khi kết thúc thu âm, anh có thể tải trực tiếp **File Ghi Âm Gốc (`.webm`)** hoặc **File Phụ Đề (`.txt`)** để đưa thẳng sang Bước 2 làm đầu vào Recap."
    )

    speech_lang = "en-US" if trans_mode == "Anh -> Việt" else "vi-VN"
    glossary_json_str = json.dumps(st.session_state.glossary, ensure_ascii=False)
    tm_json_str = json.dumps(st.session_state.tm, ensure_ascii=False)

    web_audio_component = f"""
    <div style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:10px; padding:16px; font-family:sans-serif;">
      <div style="display:flex; gap:10px; flex-wrap:wrap; align-items:center; margin-bottom:12px;">
        <select id="audioSource" style="flex:1; min-width:240px; padding:8px; border-radius:6px; border:1px solid #94a3b8; font-weight:600;"></select>
        <button id="btnStart" onclick="startCapture()" style="background:#16a34a; color:white; border:none; padding:9px 16px; border-radius:6px; font-weight:bold; cursor:pointer;">▶️ Bắt đầu Lắng nghe & Ghi âm</button>
        <button id="btnStop" onclick="stopCapture()" disabled style="background:#dc2626; color:white; border:none; padding:9px 16px; border-radius:6px; font-weight:bold; cursor:pointer; opacity:0.5;">⏹️ Dừng</button>
        <a id="downloadAudio" style="display:none; background:#2563eb; color:white; text-decoration:none; padding:9px 14px; border-radius:6px; font-weight:bold; font-size:13px;">💾 1. Tải File Ghi Âm (.webm)</a>
        <a id="downloadTxt" style="display:none; background:#0d9488; color:white; text-decoration:none; padding:9px 14px; border-radius:6px; font-weight:bold; font-size:13px;">📄 2. Tải File Phụ Đề (.txt)</a>
      </div>

      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px; font-size:13px; color:#334155;">
        <span>🎚️ <b>Tín hiệu VB-Cable (Đã bật lọc nhiễu DSP):</b></span>
        <div style="flex:1; height:10px; background:#e2e8f0; border-radius:5px; overflow:hidden;">
          <div id="volumeBar" style="width:0%; height:100%; background:#22c55e; transition:width 0.08s;"></div>
        </div>
        <span id="statusText" style="font-weight:bold; color:#64748b;">Chờ kích hoạt...</span>
      </div>

      <div id="liveSubBox" style="height:220px; overflow-y:auto; background:#0f172a; color:#f8fafc; padding:14px; border-radius:8px; font-size:15px; line-height:1.6;">
        <div id="finalSubtitles"></div>
        <div id="interimSubtitle" style="color:#38bdf8; font-style:italic;"></div>
      </div>

      <div style="margin-top:10px; display:flex; justify-content:space-between; align-items:center;">
        <span style="font-size:12px; color:#475569;">✨ Sau khi bấm <b>Dừng</b>, anh có thể tải File Ghi Âm hoặc File .TXT ở thanh trên để đưa sang Bước 2 Recap, hoặc Copy nhanh bên dưới:</span>
        <button onclick="copyTranscript()" style="background:#0284c7; color:white; border:none; padding:7px 14px; border-radius:6px; font-weight:bold; cursor:pointer;">📋 Copy toàn bộ nội dung họp</button>
      </div>
      <textarea id="fullBuffer" style="width:100%; height:85px; margin-top:6px; border-radius:6px; border:1px solid #cbd5e1; padding:8px; font-size:13px;" placeholder="Nội dung phụ đề kèm mốc thời gian sẽ tự động ghi vào đây..."></textarea>
    </div>

    <script>
      let mediaStream = null;
      let mediaRecorder = null;
      let recordedChunks = [];
      let recognition = null;
      let isCapturing = false;
      let audioContext = null;
      let analyser = null;
      const glossary = {glossary_json_str};
      const tm = {tm_json_str};

      async function loadAudioDevices() {{
        try {{
          await navigator.mediaDevices.getUserMedia({{ audio: true }});
          const devices = await navigator.mediaDevices.enumerateDevices();
          const select = document.getElementById('audioSource');
          select.innerHTML = '';
          devices.filter(d => d.kind === 'audioinput').forEach(d => {{
            const opt = document.createElement('option');
            opt.value = d.deviceId;
            opt.text = d.label || 'Microphone / Audio Input';
            if (opt.text.toLowerCase().includes('cable')) {{
              opt.selected = true;
              opt.text = '⭐ ' + opt.text + ' (Chuẩn VB-Audio Virtual Cable)';
            }}
            select.appendChild(opt);
          }});
        }} catch (e) {{
          document.getElementById('statusText').innerText = 'Vui lòng cấp quyền Micro trên trình duyệt!';
        }}
      }}
      loadAudioDevices();

      function enrichWithSharedMemory(text) {{
        let html = text;
        if (tm[text]) {{
          html += `<br><span style="color:#4ade80; font-weight:bold;">⚡ [TM Chuẩn]: ${{tm[text]}}</span>`;
        }}
        for (const [vn, en] of Object.entries(glossary)) {{
          if (vn && text.toLowerCase().includes(vn.toLowerCase())) {{
            html += ` <span style="background:#ca8a04; color:#fff; padding:1px 6px; border-radius:4px; font-size:12px;">[${{vn}} ➔ ${{en}}]</span>`;
          }} else if (en && text.toLowerCase().includes(en.toLowerCase())) {{
            html += ` <span style="background:#ca8a04; color:#fff; padding:1px 6px; border-radius:4px; font-size:12px;">[${{en}} ➔ ${{vn}}]</span>`;
          }}
        }}
        return html;
      }}

      function updateTxtDownloadLink() {{
        const textVal = document.getElementById('fullBuffer').value;
        if (!textVal.trim()) return;
        const blobTxt = new Blob([textVal], {{ type: 'text/plain;charset=utf-8' }});
        const urlTxt = URL.createObjectURL(blobTxt);
        const aTxt = document.getElementById('downloadTxt');
        aTxt.href = urlTxt;
        aTxt.download = 'Phu_De_Hop_THACO_' + new Date().toISOString().slice(0,19).replace(/:/g,'-') + '.txt';
        aTxt.style.display = 'inline-block';
      }}

      async function startCapture() {{
        const deviceId = document.getElementById('audioSource').value;
        try {{
          mediaStream = await navigator.mediaDevices.getUserMedia({{
            audio: {{
              deviceId: deviceId ? {{ exact: deviceId }} : undefined,
              noiseSuppression: true,
              echoCancellation: true,
              autoGainControl: true
            }}
          }});

          audioContext = new (window.AudioContext || window.webkitAudioContext)();
          const source = audioContext.createMediaStreamSource(mediaStream);
          analyser = audioContext.createAnalyser();
          analyser.fftSize = 256;
          source.connect(analyser);
          const dataArray = new Uint8Array(analyser.frequencyBinCount);
          function updateVolume() {{
            if (!isCapturing) return;
            analyser.getByteFrequencyData(dataArray);
            let sum = dataArray.reduce((a, b) => a + b, 0);
            let avg = Math.min(100, Math.round((sum / dataArray.length) * 1.8));
            document.getElementById('volumeBar').style.width = avg + '%';
            requestAnimationFrame(updateVolume);
          }}

          recordedChunks = [];
          mediaRecorder = new MediaRecorder(mediaStream);
          mediaRecorder.ondataavailable = e => {{ if (e.data.size > 0) recordedChunks.push(e.data); }};
          mediaRecorder.onstop = () => {{
            const blob = new Blob(recordedChunks, {{ type: 'audio/webm' }});
            const url = URL.createObjectURL(blob);
            const a = document.getElementById('downloadAudio');
            a.href = url;
            a.download = 'Ghi_Am_Hop_THACO_' + new Date().toISOString().slice(0,19).replace(/:/g,'-') + '.webm';
            a.style.display = 'inline-block';
            updateTxtDownloadLink();
          }};
          mediaRecorder.start();

          const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
          recognition = new SpeechRec();
          recognition.lang = '{speech_lang}';
          recognition.continuous = true;
          recognition.interimResults = true;

          recognition.onresult = (event) => {{
            let interim = '';
            for (let i = event.resultIndex; i < event.results.length; ++i) {{
              const transcript = event.results[i][0].transcript.trim();
              if (event.results[i].isFinal) {{
                const now = new Date().toLocaleTimeString('vi-VN');
                const decorated = enrichWithSharedMemory(transcript);
                const finalDiv = document.getElementById('finalSubtitles');
                finalDiv.innerHTML += `<div style="margin-bottom:8px; border-bottom:1px dashed #334155; padding-bottom:4px;"><span style="color:#94a3b8;">[${{now}}]</span> ${{decorated}}</div>`;
                const box = document.getElementById('liveSubBox');
                box.scrollTop = box.scrollHeight;

                const buf = document.getElementById('fullBuffer');
                buf.value += `[${{now}}] ${{transcript}}\\n`;
                buf.scrollTop = buf.scrollHeight;
                updateTxtDownloadLink();
              }} else {{
                interim += transcript;
              }}
            }}
            document.getElementById('interimSubtitle').innerText = interim ? '⏳ ' + interim : '';
          }};

          recognition.onend = () => {{ if (isCapturing) recognition.start(); }};
          isCapturing = true;
          recognition.start();
          updateVolume();

          document.getElementById('btnStart').disabled = true;
          document.getElementById('btnStart').style.opacity = '0.5';
          document.getElementById('btnStop').disabled = false;
          document.getElementById('btnStop').style.opacity = '1';
          document.getElementById('statusText').innerText = '🟢 Đang thu & lọc nhiễu...';
          document.getElementById('statusText').style.color = '#16a34a';
        }} catch (err) {{
          alert('Không thể mở thiết bị âm thanh: ' + err.message);
        }}
      }}

      function stopCapture() {{
        isCapturing = false;
        if (recognition) recognition.stop();
        if (mediaRecorder && mediaRecorder.state !== 'inactive') mediaRecorder.stop();
        if (mediaStream) mediaStream.getTracks().forEach(t => t.stop());
        if (audioContext) audioContext.close();
        updateTxtDownloadLink();
        document.getElementById('btnStart').disabled = false;
        document.getElementById('btnStart').style.opacity = '1';
        document.getElementById('btnStop').disabled = true;
        document.getElementById('btnStop').style.opacity = '0.5';
        document.getElementById('statusText').innerText = '⏹️ Đã dừng. Hãy bấm tải File Ghi Âm hoặc File .TXT!';
        document.getElementById('statusText').style.color = '#dc2626';
      }}

      function copyTranscript() {{
        const buf = document.getElementById('fullBuffer');
        buf.select();
        document.execCommand('copy');
        alert('Đã copy toàn bộ nội dung!');
      }}
    </script>
    """
    components.html(web_audio_component, height=480)

    st.markdown("#### 🔄 Bàn Dịch Thuật & Hiệu Đính Đồng Bộ (`upsert_tm`)")
    col_in, col_out = st.columns([1, 1.2])
    with col_in:
        raw_input_box = st.text_area(
            "📥 Dán nội dung cuộc họp vừa Copy (hoặc nhập câu cần dịch):",
            value=st.session_state.raw_stream_input,
            height=220,
            placeholder="[09:15:20] We need to check the warranty policy for CBU units..."
        )
        if st.button(f"🚀 Dịch Chuẩn Hóa ({selected_trans_model}) & Đẩy Lên TM", type="primary", use_container_width=True):
            st.session_state.raw_stream_input = raw_input_box
            raw_lines = [line.strip() for line in raw_input_box.split("\n") if line.strip()]
            parsed_lines = []
            for line in raw_lines:
                t_tag = datetime.now().strftime("%H:%M:%S")
                c_part = line
                if line.startswith("[") and "]" in line:
                    idx = line.find("]")
                    t_tag = line[1:idx]
                    c_part = line[idx+1:].strip()
                parsed_lines.append((t_tag, c_part))

            with st.spinner(f"Đang đối chiếu TM/Glossary và dịch bằng {selected_trans_model}..."):
                st.session_state.transcript_list = translate_batch_with_thaco_memory(
                    parsed_lines, trans_mode, gemini_api_key, selected_trans_model, webapp_url
                )
                st.success("Đã dịch và đồng bộ tự động các câu mới lên Sheet TM!")

    with col_out:
        st.markdown("**✏️ Hiệu đính nhanh bản dịch (Sửa trực tiếp cột 'Bản Dịch Chuẩn' để ghi đè lên Sheet TM):**")
        if st.session_state.transcript_list:
            edited_subtitles = st.data_editor(
                st.session_state.transcript_list,
                column_config={
                    "time": st.column_config.TextColumn("Thời gian", width="small", disabled=True),
                    "source": st.column_config.TextColumn("Câu gốc (Khoa_Goc)", width="medium"),
                    "target": st.column_config.TextColumn("Bản Dịch Chuẩn (Ban_Dich_Chuan)", width="large"),
                    "terms": st.column_config.ListColumn("Thuật ngữ áp dụng", width="small")
                },
                use_container_width=True,
                height=210,
                key="sub_editor"
            )
            st.session_state.transcript_list = edited_subtitles

            if st.button("💾 Lưu các câu vừa hiệu đính đè lên Google Sheets (`upsert_tm`)", use_container_width=True):
                upsert_dict = {item["source"]: item["target"] for item in edited_subtitles if item.get("source") and item.get("target")}
                ok, msg = upsert_tm_to_sheets(webapp_url, upsert_dict)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
        else:
            st.info("Chưa có dữ liệu dịch. Hãy bấm nút Dịch bên trái để hiển thị bảng hiệu đính.")

# ----------------- TAB 2: AI RECAP TỪ FILE GHI ÂM HOẶC FILE .TXT -----------------
with tab2:
    st.subheader(f"🧠 Tổng Hợp & Phân Tích Cuộc Họp (Model đang chọn: `{selected_recap_model}`)")

    input_source_mode = st.radio(
        "📂 Chọn nguồn đầu vào cho AI Recap:",
        [
            "🎙️ 1. Dùng File Ghi Âm (.webm / .wav / .mp3 / .m4a) tải về từ Bước 1 (AI nghe trực tiếp cuộc họp - Chuẩn xác nhất)",
            "📄 2. Dùng File Phụ Đề (.txt) tải về từ Bước 1",
            "📝 3. Dùng trực tiếp văn bản đã dịch / dán ở Bước 1"
        ]
    )

    uploaded_audio_recap = None
    uploaded_txt_recap = None
    text_for_recap = ""

    if input_source_mode.startswith("🎙️"):
        uploaded_audio_recap = st.file_uploader(
            "Tải lên File Ghi Âm cuộc họp vừa thu ở Bước 1 (.webm, .wav, .mp3, .m4a):",
            type=["webm", "wav", "mp3", "m4a", "ogg"]
        )
        if uploaded_audio_recap:
            st.audio(uploaded_audio_recap)
    elif input_source_mode.startswith("📄"):
        uploaded_txt_recap = st.file_uploader(
            "Tải lên File Phụ Đề (.txt) vừa tải xuống từ Bước 1:",
            type=["txt"]
        )
        if uploaded_txt_recap:
            text_for_recap = uploaded_txt_recap.read().decode("utf-8", errors="ignore")
            st.text_area("Nội dung trong file .txt:", value=text_for_recap, height=130)
    else:
        combined_transcript = "\n".join(
            [f"[{x['time']}] {x['target']} (Gốc: {x['source']})" for x in st.session_state.transcript_list]
        ) or st.session_state.raw_stream_input
        text_for_recap = st.text_area(
            "Nội dung văn bản cuộc họp:",
            value=combined_transcript,
            height=140
        )

    prompt_templates = {
        "1. Họp Xử lý lỗi Kỹ thuật & Chất lượng sau bán hàng (Khuyên dùng)": """Đóng vai Trưởng bộ phận Quản lý Chất lượng tại THACO AUTO. Dựa vào toàn bộ nội dung cuộc họp và bảng thuật ngữ THACO, hãy phân tích kỹ lưỡng và xuất ra định dạng JSON chuẩn với đúng các trường sau:
{
  "tom_tat_chung": "Tóm tắt mục tiêu cuộc họp, dòng xe/linh kiện liên quan và bối cảnh vấn đề (3-5 câu)",
  "noi_dung_thao_luan": "Phân tích chi tiết: (1) Hiện tượng & Số liệu thực tế, (2) Phân tích nguyên nhân gốc rễ (Root Cause), (3) Đối sách xử lý tạm thời và lâu dài",
  "ket_luan_chi_dao": "Các quyết định và chỉ đạo chốt của người chủ trì cuộc họp",
  "action_items": [
    {"tt": "1", "cong_viec": "Tên đầu việc cụ thể", "phu_trach": "Đơn vị / Cá nhân", "thoi_han": "Thời hạn hoàn thành", "ghi_chu": "Yêu cầu báo cáo"}
  ]
}
Chỉ trả về đúng chuỗi JSON hợp lệ, không thêm ký tự markdown.""",

        "2. Họp Đàm phán Đối tác Quốc tế / Nhà cung cấp (NCC)": """Đóng vai Thư ký Đối ngoại Kỹ thuật của THACO AUTO. Hãy tổng hợp nội dung làm việc với đối tác/NCC thành JSON chuẩn gồm các trường:
{
  "tom_tat_chung": "Mục đích buổi làm việc với đối tác/NCC và các nội dung chính đã trao đổi",
  "noi_dung_thao_luan": "Chi tiết các cam kết kỹ thuật, chính sách bảo hành, tiến độ cung ứng hoặc chuyển giao công nghệ",
  "ket_luan_chi_dao": "Các thỏa thuận đã thống nhất giữa hai bên và các điểm còn chờ xác nhận",
  "action_items": [
    {"tt": "1", "cong_viec": "Hạng mục cần thực hiện", "phu_trach": "THACO / Đối tác", "thoi_han": "Deadline", "ghi_chu": "Ghi chú"}
  ]
}
Chỉ trả về đúng chuỗi JSON hợp lệ, không thêm ký tự markdown.""",

        "3. Họp Giao ban Điều hành & Tiến độ Dự án": """Đóng vai Thư ký Ban Lãnh Đạo. Hãy tóm tắt cuộc họp giao ban thành JSON chuẩn gồm các trường:
{
  "tom_tat_chung": "Đánh giá tổng quan kết quả thực hiện công việc và mục tiêu trọng tâm kỳ tới",
  "noi_dung_thao_luan": "Báo cáo của từng bộ phận, các vướng mắc phát sinh và giải pháp tháo gỡ",
  "ket_luan_chi_dao": "Ý kiến chỉ đạo chốt của Lãnh đạo chủ trì",
  "action_items": [
    {"tt": "1", "cong_viec": "Nhiệm vụ trọng tâm", "phu_trach": "Phòng/Ban phụ trách", "thoi_han": "Thời hạn", "ghi_chu": "Ghi chú"}
  ]
}
Chỉ trả về đúng chuỗi JSON hợp lệ, không thêm ký tự markdown."""
    }

    selected_tpl = st.selectbox("📌 Chọn kịch bản Prompt Recap (hoặc chỉnh sửa trực tiếp bên dưới):", list(prompt_templates.keys()))
    custom_prompt = st.text_area("✍️ Nội dung Prompt giao cho AI:", value=prompt_templates[selected_tpl], height=160)

    if st.button(f"✨ Chạy AI Recap Với Model `{selected_recap_model}`", type="primary", use_container_width=True):
        if not gemini_api_key:
            st.error("Vui lòng nhập Gemini API Key ở thanh bên trái!")
        elif input_source_mode.startswith("🎙️") and not uploaded_audio_recap:
            st.warning("Vui lòng tải lên File Ghi Âm từ Bước 1!")
        elif not input_source_mode.startswith("🎙️") and not text_for_recap.strip():
            st.warning("Chưa có dữ liệu văn bản cuộc họp để phân tích!")
        else:
            with st.spinner(f"Model {selected_recap_model} đang xử lý dữ liệu cuộc họp và đối chiếu Bảng thuật ngữ THACO..."):
                try:
                    genai.configure(api_key=gemini_api_key.strip())
                    model = genai.GenerativeModel(selected_recap_model)
                    glossary_context = f"--- BẢNG THUẬT NGỮ THACO DÙNG CHUNG ---\n{json.dumps(st.session_state.glossary, ensure_ascii=False)}\n\n"

                    if input_source_mode.startswith("🎙️"):
                        # Xử lý File Ghi Âm trực tiếp qua Gemini File API
                        suffix = "." + uploaded_audio_recap.name.split(".")[-1]
                        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                            tmp_file.write(uploaded_audio_recap.getvalue())
                            tmp_path = tmp_file.name
                        
                        gemini_audio_file = genai.upload_file(path=tmp_path)
                        # Chờ file audio sẵn sàng
                        while gemini_audio_file.state.name == "PROCESSING":
                            time.sleep(1)
                            gemini_audio_file = genai.get_file(gemini_audio_file.name)

                        full_prompt = (
                            f"{custom_prompt}\n\n"
                            f"{glossary_context}"
                            f"Hãy lắng nghe toàn bộ file ghi âm cuộc họp đính kèm (có thể gồm tiếng Việt và tiếng Anh) và tổng hợp chính xác theo cấu trúc JSON đã yêu cầu."
                        )
                        res = model.generate_content([gemini_audio_file, full_prompt])
                        try:
                            genai.delete_file(gemini_audio_file.name)
                            os.remove(tmp_path)
                        except Exception:
                            pass
                    else:
                        full_prompt = (
                            f"{custom_prompt}\n\n"
                            f"{glossary_context}"
                            f"--- NỘI DUNG CUỘC HỌP ---\n{text_for_recap}"
                        )
                        res = model.generate_content(full_prompt)

                    cleaned = res.text.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                    parsed = json.loads(cleaned)
                    st.session_state.recap_data = {
                        "tom_tat_chung": parsed.get("tom_tat_chung", ""),
                        "noi_dung_thao_luan": parsed.get("noi_dung_thao_luan", ""),
                        "ket_luan_chi_dao": parsed.get("ket_luan_chi_dao", ""),
                        "action_items": parsed.get("action_items", [])
                    }
                    st.success("✅ Đã Recap thành công! Toàn bộ nội dung dưới đây đã sẵn sàng để xuất sang Biên bản Word ở Bước 3.")
                except Exception as e:
                    st.error(f"Lỗi phân tích AI: {e}")

    st.markdown("#### 🔍 Rà soát & Tinh chỉnh kết quả Recap (Liên thông trực tiếp sang Bước 3)")
    st.session_state.recap_data["tom_tat_chung"] = st.text_area(
        "1. Tóm tắt mục tiêu & bối cảnh (`{{tom_tat_chung}}`):",
        value=st.session_state.recap_data.get("tom_tat_chung", ""), height=85
    )
    st.session_state.recap_data["noi_dung_thao_luan"] = st.text_area(
        "2. Nội dung thảo luận chi tiết (`{{noi_dung_thao_luan}}`):",
        value=st.session_state.recap_data.get("noi_dung_thao_luan", ""), height=160
    )
    st.session_state.recap_data["ket_luan_chi_dao"] = st.text_area(
        "3. Kết luận & Chỉ đạo (`{{ket_luan_chi_dao}}`):",
        value=st.session_state.recap_data.get("ket_luan_chi_dao", ""), height=110
    )

    st.markdown("**4. Bảng phân công công việc (`{{ke_hoach_hanh_dong}}` / `action_items`):**")
    edited_actions = st.data_editor(
        st.session_state.recap_data.get("action_items", []) or [{"tt": "1", "cong_viec": "", "phu_trach": "", "thoi_han": "", "ghi_chu": ""}],
        num_rows="dynamic",
        use_container_width=True
    )
    st.session_state.recap_data["action_items"] = edited_actions

# ----------------- TAB 3: XUẤT BIÊN BẢN HỌP WORD TRÌNH KÝ -----------------
with tab3:
    st.subheader("📄 Xuất Biên Bản Họp Word (.docx) Trình Ký Ngay (Lấy trực tiếp từ kết quả Mục 2)")
    
    col_dl_tpl, col_guide = st.columns([1, 2])
    with col_dl_tpl:
        st.download_button(
            label="📥 Tải File Word Form Mẫu Chuẩn (.docx)",
            data=generate_sample_docx(),
            file_name="Mau_Bien_Ban_Hop_Chuan.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    with col_guide:
        st.caption(
            "Các trường tự động đổ từ Mục 2 vào File Word mẫu của anh: "
            "`{{ten_cuoc_hop}}`, `{{ngay_hop}}`, `{{dia_diem}}`, `{{chu_tri}}`, `{{thanh_phan}}`, `{{nguoi_lap}}`, "
            "`{{tom_tat_chung}}`, `{{noi_dung_thao_luan}}`, `{{ket_luan_chi_dao}}`, `{{ke_hoach_hanh_dong}}` (và bảng lặp `action_items`)."
        )

    c1, c2, c3 = st.columns(3)
    with c1:
        ten_hop = st.text_input("Tên cuộc họp:", value="Họp đánh giá chất lượng & xử lý kỹ thuật")
        ngay_hop = st.text_input("Thời gian họp:", value=datetime.now().strftime("%H:%M ngày %d/%m/%Y"))
    with c2:
        dia_diem = st.text_input("Địa điểm / Hình thức:", value="Trực tuyến qua Microsoft Teams")
        chu_tri = st.text_input("Chủ trì cuộc họp:", value="")
    with c3:
        nguoi_lap = st.text_input("Người lập biên bản:", value="")
        thanh_phan = st.text_input("Thành phần tham dự:", value="Các bộ phận liên quan")

    uploaded_docx = st.file_uploader("📂 Tải lên File Word Biên Bản Mẫu (.docx) của đơn vị:", type=["docx"])

    # Hiển thị bản xem trước dữ liệu từ Bước 2 sẽ đưa vào Word
    with st.expander("👁️ Xem trước dữ liệu Recap từ Bước 2 sẽ điền vào Biên bản Word", expanded=True):
        st.markdown(f"- **Tóm tắt chung:** {st.session_state.recap_data.get('tom_tat_chung') or '*(Chưa có dữ liệu từ Bước 2)*'}")
        st.markdown(f"- **Nội dung thảo luận:** {st.session_state.recap_data.get('noi_dung_thao_luan') or '*(Chưa có dữ liệu từ Bước 2)*'}")
        st.markdown(f"- **Kết luận chỉ đạo:** {st.session_state.recap_data.get('ket_luan_chi_dao') or '*(Chưa có dữ liệu từ Bước 2)*'}")

    if st.button("🖨️ Kết Xuất Biên Bản Word Hoàn Chỉnh Để Trình Ký", type="primary", use_container_width=True):
        try:
            template_stream = uploaded_docx if uploaded_docx else io.BytesIO(generate_sample_docx())
            doc = DocxTemplate(template_stream)

            action_lines = []
            for idx, item in enumerate(st.session_state.recap_data.get("action_items", []), 1):
                if item.get("cong_viec"):
                    action_lines.append(
                        f"{idx}. {item.get('cong_viec', '')} | Phụ trách: {item.get('phu_trach', '')} "
                        f"| Thời hạn: {item.get('thoi_han', '')} ({item.get('ghi_chu', '')})"
                    )
            action_str = "\n".join(action_lines) if action_lines else "Thực hiện theo kết luận chỉ đạo."

            context = {
                "ten_cuoc_hop": ten_hop,
                "ngay_hop": ngay_hop,
                "dia_diem": dia_diem,
                "chu_tri": chu_tri,
                "thanh_phan": thanh_phan,
                "nguoi_lap": nguoi_lap,
                "tom_tat_chung": st.session_state.recap_data.get("tom_tat_chung", ""),
                "noi_dung_thao_luan": st.session_state.recap_data.get("noi_dung_thao_luan", ""),
                "ket_luan_chi_dao": st.session_state.recap_data.get("ket_luan_chi_dao", ""),
                "ke_hoach_hanh_dong": action_str,
                "action_items": st.session_state.recap_data.get("action_items", [])
            }

            doc.render(context)
            out_bio = io.BytesIO()
            doc.save(out_bio)
            out_bio.seek(0)

            out_name = f"Bien_Ban_Hop_THACO_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
            st.download_button(
                label=f"⬇️ TẢI XUỐNG BIÊN BẢN TRÌNH KÝ ({out_name})",
                data=out_bio.getvalue(),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
            st.success("✅ Đã tạo biên bản thành công từ nội dung Recap Mục 2!")
        except Exception as e:
            st.error(f"Lỗi khi xuất file Word: {e}")

# ----------------- TAB 4: QUẢN TRỊ & HIỆU ĐÍNH TRỰC TIẾP KHO CHUNG -----------------
with tab4:
    st.subheader("📚 Quản Trị & Hiệu Đính Trực Tiếp Kho Dữ Liệu Dùng Chung (Google Sheets)")
    col_g, col_tm = st.columns(2)
    
    with col_g:
        st.markdown(f"**📖 Bảng Thuật Ngữ Chung (`Glossary`: {len(st.session_state.glossary)} từ)**")
        g_list = [{"Tu_Goc": k, "Tu_Dich": v} for k, v in st.session_state.glossary.items()] or [{"Tu_Goc": "", "Tu_Dich": ""}]
        edited_g = st.data_editor(g_list, num_rows="dynamic", use_container_width=True, height=340, key="editor_glossary")
        
        if st.button("💾 Lưu toàn bộ Bảng Thuật Ngữ lên Sheets (`save_glossary`)", use_container_width=True):
            new_g_dict = {row["Tu_Goc"].strip(): row["Tu_Dich"].strip() for row in edited_g if row.get("Tu_Goc") and row.get("Tu_Dich")}
            ok, msg = save_full_glossary_to_sheets(webapp_url, new_g_dict)
            if ok:
                st.success(msg)
            else:
                st.error(msg)

    with col_tm:
        st.markdown(f"**🧠 Bộ Nhớ Đệm Dịch Thuật Chung (`TM`: {len(st.session_state.tm)} câu)**")
        tm_list = [{"Khoa_Goc": k, "Ban_Dich_Chuan": v} for k, v in st.session_state.tm.items()] or [{"Khoa_Goc": "", "Ban_Dich_Chuan": ""}]
        edited_tm = st.data_editor(tm_list, num_rows="dynamic", use_container_width=True, height=340, key="editor_tm")
        
        if st.button("💾 Cập nhật các câu đã sửa lên Sheets (`upsert_tm`)", use_container_width=True):
            new_tm_dict = {row["Khoa_Goc"].strip(): row["Ban_Dich_Chuan"].strip() for row in edited_tm if row.get("Khoa_Goc") and row.get("Ban_Dich_Chuan")}
            ok, msg = upsert_tm_to_sheets(webapp_url, new_tm_dict)
            if ok:
                st.success(msg)
            else:
                st.error(msg)