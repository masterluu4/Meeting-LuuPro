import streamlit as st
import streamlit.components.v1 as components
import requests
import json
import io
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
    "recap_data": {
        "tom_tat_chung": "",
        "noi_dung_thao_luan": "",
        "ket_luan_chi_dao": "",
        "action_items": []
    }
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ================= HÀM KẾT NỐI CHUẨN VỚI CODE.GS HIỆN TẠI CỦA ANH =================
def sync_from_sheets(webapp_url):
    """Đọc dữ liệu từ doGet(e) trả về { glossary: {}, tm: {} }"""
    if not webapp_url:
        return False, "⚠️ Chưa nhập Link Web App URL."
    try:
        res = requests.get(webapp_url.strip(), timeout=12)
        if res.status_code == 200:
            data = res.json()
            st.session_state.glossary = data.get("glossary", {})
            st.session_state.tm = data.get("tm", {})
            st.session_state.is_synced = True
            return True, f"✅ Đã đồng bộ chung: {len(st.session_state.glossary)} thuật ngữ & {len(st.session_state.tm)} câu TM."
    except Exception as e:
        return False, f"❌ Lỗi kết nối Google Sheets: {e}"
    return False, "❌ Đường dẫn Web App URL không phản hồi."

def save_full_glossary_to_sheets(webapp_url, updated_glossary_dict):
    """
    Khớp với action === 'save_glossary' trong Code.gs:
    Vì Code.gs dùng gSheet.clearContents(), hàm này luôn kéo dữ liệu mới nhất trên mây về,
    gộp với dữ liệu mới rồi mới đẩy toàn bộ lên để bảo vệ dữ liệu của phần mềm dịch tài liệu.
    """
    if not webapp_url:
        return False, "Chưa nhập Link Web App URL."
    try:
        # Bước 1: Lấy dữ liệu mới nhất từ Sheets để tránh ghi đè mất từ của người khác
        res_get = requests.get(webapp_url.strip(), timeout=10)
        cloud_glossary = {}
        if res_get.status_code == 200:
            cloud_glossary = res_get.json().get("glossary", {})
        
        # Bước 2: Gộp từ điển trên mây và từ điển cập nhật
        cloud_glossary.update(updated_glossary_dict)
        
        # Bước 3: Đẩy toàn bộ lên theo đúng cấu trúc payload.action = "save_glossary"
        payload = {
            "action": "save_glossary",
            "data": cloud_glossary
        }
        res_post = requests.post(webapp_url.strip(), data=json.dumps(payload), timeout=12)
        if res_post.status_code == 200 and res_post.json().get("status") == "ok":
            st.session_state.glossary = cloud_glossary
            return True, f"✅ Đã lưu đồng bộ {len(cloud_glossary)} thuật ngữ lên Google Sheets!"
        return False, f"❌ Lỗi từ Apps Script: {res_post.text}"
    except Exception as e:
        return False, f"❌ Lỗi kết nối: {e}"

def upsert_tm_to_sheets(webapp_url, new_tm_dict):
    """
    Khớp với action === 'upsert_tm' trong Code.gs:
    Hỗ trợ đẩy 1 câu hoặc hàng loạt câu cùng lúc. Nếu câu đã có sẽ tự động ghi đè bản dịch chuẩn mới.
    """
    if not webapp_url or not new_tm_dict:
        return False, "Không có dữ liệu hoặc thiếu URL."
    try:
        clean_dict = {str(k).strip(): str(v).strip() for k, v in new_tm_dict.items() if str(k).strip() and str(v).strip()}
        payload = {
            "action": "upsert_tm",
            "data": clean_dict
        }
        res = requests.post(webapp_url.strip(), data=json.dumps(payload), timeout=12)
        if res.status_code == 200 and res.json().get("status") == "ok":
            st.session_state.tm.update(clean_dict)
            return True, f"✅ Đã cập nhật {len(clean_dict)} câu vào Bộ nhớ đệm (TM) dùng chung!"
        return False, f"❌ Lỗi từ Apps Script: {res.text}"
    except Exception as e:
        return False, f"❌ Lỗi kết nối: {e}"

# ================= HÀM DỊCH THUẬT AI + SO KHỚP GLOSSARY & TM =================
def translate_batch_with_thaco_memory(lines_with_time, mode, api_key, webapp_url=""):
    """Dịch thông minh kết hợp TM (0ms), Glossary và tự động gom nhóm đẩy lên Sheets 1 lần"""
    results = []
    tm_to_upsert = {}

    for time_tag, clean_text in lines_with_time:
        if not clean_text:
            continue
        if mode == "Việt -> Việt (Chỉ ghi phụ đề)":
            results.append({"time": time_tag, "source": clean_text, "target": clean_text, "terms": []})
            continue

        # 1. Tra cứu trực tiếp trong Bộ nhớ đệm (TM) dùng chung
        if clean_text in st.session_state.tm:
            results.append({
                "time": time_tag,
                "source": clean_text,
                "target": st.session_state.tm[clean_text],
                "terms": ["⚡ TM Dùng Chung"]
            })
            continue

        # 2. Quét các thuật ngữ trong Glossary xuất hiện trong câu nói
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
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-1.5-flash")
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

            # Đưa vào danh sách chờ đẩy lên TM dùng chung (chỉ lưu câu có ý nghĩa từ 3 từ trở lên)
            if len(clean_text.split()) >= 3:
                tm_to_upsert[clean_text] = translated

        except Exception as e:
            results.append({"time": time_tag, "source": clean_text, "target": f"[Lỗi AI: {e}]", "terms": []})

    # 3. Đẩy toàn bộ câu mới dịch lên Google Sheets TM trong 1 lượt gọi duy nhất (Batch Upsert)
    if webapp_url and tm_to_upsert:
        upsert_tm_to_sheets(webapp_url, tm_to_upsert)

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
    p_info.add_run("3. Địa điểm / Nền tảng: ").bold = True
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
    
    # Tự động đọc từ Streamlit Secrets nếu có cài đặt sẵn
    default_key = st.secrets.get("GEMINI_API_KEY", "") if "GEMINI_API_KEY" in st.secrets else ""
    default_url = st.secrets.get("WEBAPP_URL", "") if "WEBAPP_URL" in st.secrets else ""

    gemini_api_key = st.text_input("🔑 1. Gemini API Key:", value=default_key, type="password")
    webapp_url = st.text_input("☁️ 2. Link Google Apps Script (Dùng chung):", value=default_url)
    
    # Tự động tải Glossary & TM ngay lần đầu mở ứng dụng nếu đã có URL
    if webapp_url and not st.session_state.is_synced:
        ok, msg = sync_from_sheets(webapp_url)
        if ok:
            st.toast(msg)

    if st.button("🔄 Đồng bộ lại Glossary & TM từ Sheets", use_container_width=True, type="primary"):
        ok, msg = sync_from_sheets(webapp_url)
        st.success(msg) if ok else st.error(msg)

    st.caption(f"📊 Trạng thái Kho dữ liệu chung: **{len(st.session_state.glossary)}** thuật ngữ | **{len(st.session_state.tm)}** câu TM")

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
                st.success(msg) if ok else st.error(msg)
            else:
                st.warning("Vui lòng nhập đủ Từ gốc, Từ dịch và Link Web App URL.")

# ================= GIAO DIỆN CHÍNH =================
st.title("🚗 Hệ Thống Trợ Lý Họp, Dịch Thuật Real-time & Xuất Biên Bản THACO")

tab1, tab2, tab3, tab4 = st.tabs([
    "🎙️ 1. Thu Âm VB-Cable & Phụ Đề Real-time",
    "🧠 2. AI Recap Cuộc Họp (Đa Kịch Bản)",
    "📄 3. Xuất Biên Bản Word Trình Ký",
    "📚 4. Quản Trị & Hiệu Đính Kho TM / Glossary Chung"
])

# ----------------- TAB 1: THU ÂM VB-CABLE TRÊN WEB & PHỤ ĐỀ -----------------
with tab1:
    st.info(
        "🔊 **Thiết lập chuẩn với VB-Audio Virtual Cable:** "
        "Trong cài đặt âm thanh của **MS Teams / Zoom / VooV**, chọn Loa (Speaker) là `CABLE Input`. "
        "Ở ô chọn thiết bị bên dưới, hệ thống sẽ tự động nhận diện và chọn **`⭐ CABLE Output (VB-Audio Virtual Cable)`**."
    )

    speech_lang = "en-US" if trans_mode == "Anh -> Việt" else "vi-VN"
    glossary_json_str = json.dumps(st.session_state.glossary, ensure_ascii=False)
    tm_json_str = json.dumps(st.session_state.tm, ensure_ascii=False)

    web_audio_component = f"""
    <div style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:10px; padding:16px; font-family:sans-serif;">
      <div style="display:flex; gap:10px; flex-wrap:wrap; align-items:center; margin-bottom:12px;">
        <select id="audioSource" style="flex:1; min-width:260px; padding:8px; border-radius:6px; border:1px solid #94a3b8; font-weight:600;"></select>
        <button id="btnStart" onclick="startCapture()" style="background:#16a34a; color:white; border:none; padding:9px 16px; border-radius:6px; font-weight:bold; cursor:pointer;">▶️ Bắt đầu Lắng nghe & Ghi âm</button>
        <button id="btnStop" onclick="stopCapture()" disabled style="background:#dc2626; color:white; border:none; padding:9px 16px; border-radius:6px; font-weight:bold; cursor:pointer; opacity:0.5;">⏹️ Dừng</button>
        <a id="downloadAudio" style="display:none; background:#2563eb; color:white; text-decoration:none; padding:9px 14px; border-radius:6px; font-weight:bold; font-size:13px;">💾 Tải File Ghi Âm Gốc</a>
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
        <span style="font-size:12px; color:#475569;">✨ Câu nào trùng khớp trong <b>TM dùng chung</b> sẽ hiển thị bản dịch ngay lập tức. Bấm Copy để chuyển xuống bàn dịch AI & hiệu đính:</span>
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
        // 1. Nếu câu khớp hoàn toàn trong TM dùng chung -> Hiển thị bản dịch chuẩn tức thời
        if (tm[text]) {{
          html += `<br><span style="color:#4ade80; font-weight:bold;">⚡ [TM Chuẩn]: ${{tm[text]}}</span>`;
        }}
        // 2. Tô sáng các thuật ngữ trong Glossary dùng chung
        for (const [vn, en] of Object.entries(glossary)) {{
          if (vn && text.toLowerCase().includes(vn.toLowerCase())) {{
            html += ` <span style="background:#ca8a04; color:#fff; padding:1px 6px; border-radius:4px; font-size:12px;">[${{vn}} ➔ ${{en}}]</span>`;
          }} else if (en && text.toLowerCase().includes(en.toLowerCase())) {{
            html += ` <span style="background:#ca8a04; color:#fff; padding:1px 6px; border-radius:4px; font-size:12px;">[${{en}} ➔ ${{vn}}]</span>`;
          }}
        }}
        return html;
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
        document.getElementById('btnStart').disabled = false;
        document.getElementById('btnStart').style.opacity = '1';
        document.getElementById('btnStop').disabled = true;
        document.getElementById('btnStop').style.opacity = '0.5';
        document.getElementById('statusText').innerText = '⏹️ Đã dừng thu âm.';
        document.getElementById('statusText').style.color = '#dc2626';
      }}

      function copyTranscript() {{
        const buf = document.getElementById('fullBuffer');
        buf.select();
        document.execCommand('copy');
        alert('Đã copy toàn bộ nội dung! Hãy dán vào ô bên dưới để Dịch chuẩn hóa & Tự động cập nhật vào TM dùng chung.');
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
            height=230,
            placeholder="[09:15:20] We need to check the warranty policy for CBU units..."
        )
        if st.button("🚀 Dịch Chuẩn Hóa & Đẩy Lên TM Dùng Chung", type="primary", use_container_width=True):
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

            with st.spinner("Đang đối chiếu TM/Glossary dùng chung và dịch bằng Gemini Flash..."):
                st.session_state.transcript_list = translate_batch_with_thaco_memory(
                    parsed_lines, trans_mode, gemini_api_key, webapp_url
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
                st.success(msg) if ok else st.error(msg)
        else:
            st.info("Chưa có dữ liệu dịch. Hãy bấm nút Dịch bên trái để hiển thị bảng hiệu đính.")

# ----------------- TAB 2: AI RECAP CUỘC HỌP (ĐA KỊCH BẢN) -----------------
with tab2:
    st.subheader("🧠 Tổng Hợp & Phân Tích Cuộc Họp Bằng Gemini Pro")

    prompt_templates = {
        "1. Họp Xử lý lỗi Kỹ thuật & Chất lượng sau bán hàng (Khuyên dùng)": """Đóng vai Trưởng bộ phận Quản lý Chất lượng tại THACO AUTO. Dựa vào nội dung cuộc họp và bảng thuật ngữ THACO, hãy phân tích và xuất ra định dạng JSON chuẩn với các trường sau:
{
  "tom_tat_chung": "Tóm tắt mục tiêu cuộc họp, dòng xe/linh kiện liên quan và bối cảnh vấn đề (3-5 câu)",
  "noi_dung_thao_luan": "Phân tích rõ: (1) Hiện tượng & Dữ liệu thực tế, (2) Phân tích nguyên nhân gốc rễ (Root Cause), (3) Đối sách xử lý tạm thời và lâu dài",
  "ket_luan_chi_dao": "Các quyết định chốt của người chủ trì cuộc họp",
  "action_items": [
    {"cong_viec": "Tên đầu việc cụ thể", "phu_trach": "Đơn vị / Cá nhân", "thoi_han": "Thời hạn hoàn thành", "ghi_chu": "Yêu cầu báo cáo"}
  ]
}
Chỉ trả về đúng chuỗi JSON hợp lệ, không thêm ký tự markdown.""",

        "2. Họp Đàm phán Đối tác Quốc tế / Nhà cung cấp (NCC)": """Đóng vai Thư ký Đối ngoại Kỹ thuật của THACO AUTO. Hãy tổng hợp nội dung làm việc với đối tác/NCC thành JSON chuẩn gồm các trường:
{
  "tom_tat_chung": "Mục đích buổi làm việc với đối tác/NCC và các nội dung chính đã trao đổi",
  "noi_dung_thao_luan": "Chi tiết các cam kết kỹ thuật, chính sách bảo hành, tiến độ cung ứng hoặc chuyển giao công nghệ",
  "ket_luan_chi_dao": "Các thỏa thuận đã thống nhất giữa hai bên và các điểm còn chờ xác nhận",
  "action_items": [
    {"cong_viec": "Hạng mục cần thực hiện", "phu_trach": "THACO / Đối tác", "thoi_han": "Deadline", "ghi_chu": "Ghi chú"}
  ]
}
Chỉ trả về đúng chuỗi JSON hợp lệ, không thêm ký tự markdown.""",

        "3. Họp Giao ban Điều hành & Tiến độ Dự án": """Đóng vai Thư ký Ban Lãnh Đạo. Hãy tóm tắt cuộc họp giao ban thành JSON chuẩn gồm các trường:
{
  "tom_tat_chung": "Đánh giá tổng quan kết quả thực hiện công việc và mục tiêu trọng tâm kỳ tới",
  "noi_dung_thao_luan": "Báo cáo của từng bộ phận, các vướng mắc phát sinh và giải pháp tháo gỡ",
  "ket_luan_chi_dao": "Ý kiến chỉ đạo chốt của Lãnh đạo chủ trì",
  "action_items": [
    {"cong_viec": "Nhiệm vụ trọng tâm", "phu_trach": "Phòng/Ban phụ trách", "thoi_han": "Thời hạn", "ghi_chu": "Ghi chú"}
  ]
}
Chỉ trả về đúng chuỗi JSON hợp lệ, không thêm ký tự markdown."""
    }

    selected_tpl = st.selectbox("📂 Chọn kịch bản Prompt mẫu (hoặc tùy chỉnh trực tiếp bên dưới):", list(prompt_templates.keys()))
    custom_prompt = st.text_area("✍️ Nội dung Prompt giao cho Gemini Pro:", value=prompt_templates[selected_tpl], height=170)

    combined_transcript = "\n".join(
        [f"[{x['time']}] {x['target']} (Gốc: {x['source']})" for x in st.session_state.transcript_list]
    ) or st.session_state.raw_stream_input

    meeting_text_for_recap = st.text_area(
        "📝 Toàn bộ nội dung cuộc họp đưa vào phân tích:",
        value=combined_transcript,
        height=150
    )

    if st.button("✨ Chạy AI Gemini Pro Tổng Hợp Biên Bản", type="primary", use_container_width=True):
        if not gemini_api_key:
            st.error("Vui lòng nhập Gemini API Key ở thanh bên trái!")
        elif not meeting_text_for_recap.strip():
            st.warning("Chưa có dữ liệu cuộc họp để phân tích!")
        else:
            with st.spinner("Gemini Pro đang phân tích chuyên sâu cùng Bảng thuật ngữ THACO..."):
                try:
                    genai.configure(api_key=gemini_api_key)
                    model = genai.GenerativeModel("gemini-1.5-pro")
                    full_prompt = (
                        f"{custom_prompt}\n\n"
                        f"--- BẢNG THUẬT NGỮ THACO DÙNG CHUNG ---\n{json.dumps(st.session_state.glossary, ensure_ascii=False)}\n\n"
                        f"--- NỘI DUNG CUỘC HỌP ---\n{meeting_text_for_recap}"
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
                    st.success("✅ Đã phân tích xong! Anh có thể rà soát và tinh chỉnh ngay bên dưới trước khi xuất Word.")
                except Exception as e:
                    st.error(f"Lỗi phân tích AI: {e}")

    st.markdown("#### 🔍 Rà soát & Tinh chỉnh nội dung Biên bản")
    st.session_state.recap_data["tom_tat_chung"] = st.text_area(
        "1. Tóm tắt mục tiêu & bối cảnh (`{{tom_tat_chung}}`):",
        value=st.session_state.recap_data.get("tom_tat_chung", ""), height=90
    )
    st.session_state.recap_data["noi_dung_thao_luan"] = st.text_area(
        "2. Nội dung thảo luận chi tiết (`{{noi_dung_thao_luan}}`):",
        value=st.session_state.recap_data.get("noi_dung_thao_luan", ""), height=160
    )
    st.session_state.recap_data["ket_luan_chi_dao"] = st.text_area(
        "3. Kết luận & Chỉ đạo (`{{ket_luan_chi_dao}}`):",
        value=st.session_state.recap_data.get("ket_luan_chi_dao", ""), height=110
    )

    st.markdown("**4. Bảng phân công công việc (`{{ke_hoach_hanh_dong}}`):**")
    edited_actions = st.data_editor(
        st.session_state.recap_data.get("action_items", []) or [{"cong_viec": "", "phu_trach": "", "thoi_han": "", "ghi_chu": ""}],
        num_rows="dynamic",
        use_container_width=True
    )
    st.session_state.recap_data["action_items"] = edited_actions

# ----------------- TAB 3: XUẤT BIÊN BẢN HỌP WORD TRÌNH KÝ -----------------
with tab3:
    st.subheader("📄 Xuất Biên Bản Họp Word (.docx) Trình Ký Ngay")
    
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
            "Các biến tự động điền trong file Word mẫu: "
            "`{{ten_cuoc_hop}}`, `{{ngay_hop}}`, `{{dia_diem}}`, `{{chu_tri}}`, `{{thanh_phan}}`, `{{nguoi_lap}}`, "
            "`{{tom_tat_chung}}`, `{{noi_dung_thao_luan}}`, `{{ket_luan_chi_dao}}`, `{{ke_hoach_hanh_dong}}`."
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

    uploaded_docx = st.file_uploader("📂 Tải lên File Word Form Mẫu của đơn vị (.docx):", type=["docx"])

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
            st.success("✅ Đã tạo biên bản thành công!")
        except Exception as e:
            st.error(f"Lỗi khi xuất file Word: {e}")

# ----------------- TAB 4: QUẢN TRỊ & HIỆU ĐÍNH TRỰC TIẾP KHO CHUNG -----------------
with tab4:
    st.subheader("📚 Quản Trị & Hiệu Đính Trực Tiếp Kho Dữ Liệu Dùng Chung (Google Sheets)")
    st.caption("Mọi chỉnh sửa tại đây sẽ đồng bộ 2 chiều theo đúng chuẩn `save_glossary` và `upsert_tm` với phần mềm dịch tài liệu của anh.")

    col_g, col_tm = st.columns(2)
    
    with col_g:
        st.markdown(f"**📖 Bảng Thuật Ngữ Chung (`Glossary`: {len(st.session_state.glossary)} từ)**")
        g_list = [{"Tu_Goc": k, "Tu_Dich": v} for k, v in st.session_state.glossary.items()]
        if not g_list:
            g_list = [{"Tu_Goc": "", "Tu_Dich": ""}]
        
        edited_g = st.data_editor(g_list, num_rows="dynamic", use_container_width=True, height=340, key="editor_glossary")
        
        if st.button("💾 Lưu toàn bộ Bảng Thuật Ngữ lên Sheets (`save_glossary`)", use_container_width=True):
            new_g_dict = {row["Tu_Goc"].strip(): row["Tu_Dich"].strip() for row in edited_g if row.get("Tu_Goc") and row.get("Tu_Dich")}
            ok, msg = save_full_glossary_to_sheets(webapp_url, new_g_dict)
            st.success(msg) if ok else st.error(msg)

    with col_tm:
        st.markdown(f"**🧠 Bộ Nhớ Đệm Dịch Thuật Chung (`TM`: {len(st.session_state.tm)} câu)**")
        tm_list = [{"Khoa_Goc": k, "Ban_Dich_Chuan": v} for k, v in st.session_state.tm.items()]
        if not tm_list:
            tm_list = [{"Khoa_Goc": "", "Ban_Dich_Chuan": ""}]
            
        edited_tm = st.data_editor(tm_list, num_rows="dynamic", use_container_width=True, height=340, key="editor_tm")
        
        if st.button("💾 Cập nhật các câu đã sửa lên Sheets (`upsert_tm`)", use_container_width=True):
            new_tm_dict = {row["Khoa_Goc"].strip(): row["Ban_Dich_Chuan"].strip() for row in edited_tm if row.get("Khoa_Goc") and row.get("Ban_Dich_Chuan")}
            ok, msg = upsert_tm_to_sheets(webapp_url, new_tm_dict)
            st.success(msg) if ok else st.error(msg)