import streamlit as st
import streamlit.components.v1 as components
import requests
import json
import io
import re
import time
from datetime import datetime
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ================= CẤU HÌNH TRANG =================
st.set_page_config(page_title="THACO Smart Meeting & AI Biên Bản - Lưu Pro", page_icon="🚗", layout="wide")

DEFAULT_PERSONNEL = [
    {"ho_ten": "Ông Nguyễn Thiện Đức", "chuc_danh": "PGĐ.QLCL Ô tô & LKPT KCN", "phong_ban": "Phòng QLCL KCN"},
    {"ho_ten": "Ông Võ Thành Lưu", "chuc_danh": "TP QLCL xe Thương mại", "phong_ban": "Phòng QLCL KCN"},
    {"ho_ten": "Ông Nguyễn Ngọc Sinh", "chuc_danh": "TBP KTCL NM Bus", "phong_ban": "Nhà máy Tải & Bus"},
    {"ho_ten": "Ông Bùi Quang Trung", "chuc_danh": "TBP KTCL NM Tải", "phong_ban": "Nhà máy Tải & Bus"},
    {"ho_ten": "Ông Trần Đình Quang", "chuc_danh": "TP Khối KD Xuất khẩu", "phong_ban": "Phòng KD Xuất khẩu"}
]

# Khởi tạo Session State
defaults = {
    "transcript_list": [],
    "raw_stream_input": "",
    "glossary": {},
    "tm": {},
    "personnel": DEFAULT_PERSONNEL,
    "selected_attendees": [],        # Danh sách nhân sự tham gia được chọn ở Bước 2
    "extra_attendees_str": "",       # Nhân sự / Đối tác bổ sung ngoài danh bạ
    "template_docx_bytes": None,     # Lưu trữ file Word BBH mẫu tải lên từ Bước 2
    "template_docx_name": "",        # Tên file Word BBH mẫu
    "template_structure_hint": "",   # Cấu trúc trích xuất từ file BBH mẫu để AI hiểu form
    "is_synced": False,
    "bbh_meta": {
        "so_bbh": "…/2026/BBH – Phòng QLCL KCN",
        "ve_viec": "Đánh giá chất lượng xe và giải quyết các vấn đề kỹ thuật sau bán hàng",
        "thoi_gian": datetime.now().strftime("%H:%M ngày %d/%m/%Y"),
        "dia_diem": "Họp trực tuyến qua Microsoft Teams / Phòng họp QLCL KCN",
        "ket_luan_chung": "Trên đây là các thống nhất, chỉ đạo và triển khai trong cuộc họp, đề nghị các phòng ban, các Nhà máy triển khai thực hiện theo tiến độ, Phòng QLCL KCN làm đầu mối tổng hợp, theo dõi và báo cáo tiến độ thực hiện theo định kỳ."
    },
    "recap_rows": [
        {
            "tt": "1",
            "van_de": "Đối với 04 báo cáo khiếu nại chưa có Báo cáo kỹ thuật từ BJM (chỉ có thông tin nhanh)",
            "noi_dung_thao_luan": "BJM xác nhận xe của khách hàng cuối đang vận hành ở xa (không vào xưởng) nên chưa thể chụp hình và lập báo cáo kỹ thuật.",
            "cong_viec_tiep_theo": "BJM sẽ theo dõi tình hình và cập nhật thông tin cho THACO trong tuần tiếp theo.",
            "trach_nhiem": "BJM",
            "phoi_hop": "-",
            "hoan_thanh": "20/11/2026"
        }
    ]
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ================= GIAO DIỆN CSS COCKPIT THACO EXECUTIVE =================
st.markdown("""
<style>
    .block-container { padding-top: 1.1rem !important; padding-bottom: 1rem !important; max-width: 98% !important; }
    
    .thaco-header {
        background: linear-gradient(90deg, #002B5B 0%, #004B8D 100%);
        color: white; padding: 12px 22px; border-radius: 10px;
        display: flex; justify-content: space-between; align-items: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.12); margin-bottom: 12px;
        flex-wrap: wrap; gap: 10px;
    }
    .thaco-title { font-size: 19px; font-weight: 700; letter-spacing: 0.4px; margin: 0; }
    .thaco-badges { display: flex; gap: 8px; font-size: 12.5px; flex-wrap: wrap; }
    .thaco-badge {
        background: rgba(255,255,255,0.16); padding: 4px 12px;
        border-radius: 20px; border: 1px solid rgba(255,255,255,0.28); font-weight: 600;
    }

    .stTabs [data-baseweb="tab-list"] { gap: 8px; background-color: #f1f5f9; padding: 6px; border-radius: 10px; }
    .stTabs [data-baseweb="tab"] {
        height: 40px; border-radius: 8px; padding: 0 16px;
        font-weight: 600; color: #334155; background-color: transparent;
    }
    .stTabs [aria-selected="true"] {
        background-color: #003366 !important; color: #ffffff !important;
        box-shadow: 0 2px 6px rgba(0,51,102,0.25);
    }

    .step-card {
        background: #f8fafc; border: 1px solid #e2e8f0; border-left: 4px solid #003366;
        border-radius: 8px; padding: 12px 16px; margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)

# ================= HÀM LÀM SẠCH API KEY & URL =================
def clean_api_key(raw_key):
    if not raw_key:
        return ""
    return str(raw_key).strip().strip("[]\"'() ")

def clean_webapp_url(raw_url):
    if not raw_url:
        return ""
    s = str(raw_url).strip()
    match = re.search(r'https?://[^\s\[\]"\'<>()]+', s)
    if match:
        return match.group(0)
    return s.strip("[]\"'() ")

# ================= HÀM ĐỒNG BỘ GOOGLE SHEETS (GLOSSARY + TM + PERSONNEL) =================
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
            cloud_personnel = data.get("personnel", [])
            if cloud_personnel:
                st.session_state.personnel = cloud_personnel
            st.session_state.is_synced = True
            return True, (
                f"✅ Đã đồng bộ: {len(st.session_state.glossary)} thuật ngữ, "
                f"{len(st.session_state.tm)} câu TM & {len(st.session_state.personnel)} nhân sự."
            )
    except Exception as e:
        return False, f"❌ Lỗi kết nối Google Sheets: {e}"
    return False, "❌ Đường dẫn không phản hồi."

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
            return True, f"✅ Đã cập nhật {len(clean_dict)} câu vào Bộ nhớ đệm (TM)!"
        return False, f"❌ Lỗi từ Apps Script: {res.text}"
    except Exception as e:
        return False, f"❌ Lỗi kết nối: {e}"

def save_personnel_to_sheets(raw_url, personnel_list):
    webapp_url = clean_webapp_url(raw_url)
    if not webapp_url:
        return False, "Chưa nhập Link Web App URL."
    try:
        clean_list = []
        for p in personnel_list:
            if p.get("ho_ten") and str(p["ho_ten"]).strip():
                clean_list.append({
                    "ho_ten": str(p.get("ho_ten", "")).strip(),
                    "chuc_danh": str(p.get("chuc_danh", "")).strip(),
                    "phong_ban": str(p.get("phong_ban", "")).strip()
                })
        payload = {"action": "save_personnel", "data": clean_list}
        res = requests.post(webapp_url, data=json.dumps(payload), timeout=12)
        if res.status_code == 200 and res.json().get("status") == "ok":
            st.session_state.personnel = clean_list
            return True, f"✅ Đã lưu {len(clean_list)} nhân sự lên tab Personnel của Google Sheets!"
        return False, f"❌ Lỗi từ Apps Script: {res.text}"
    except Exception as e:
        return False, f"❌ Lỗi kết nối: {e}"

# ================= HÀM ĐỌC CẤU TRÚC FILE WORD MẪU ĐỂ ĐƯA VÀO AI RECAP =================
def inspect_docx_template(file_bytes):
    """
    Đọc cấu trúc thực tế của file Word BBH mẫu (các đoạn tiêu đề, cấu trúc bảng, dòng ví dụ mẫu)
    để truyền cho AI Recap hiểu chính xác form mẫu và văn phong của đơn vị.
    """
    try:
        doc = Document(io.BytesIO(file_bytes))
        paragraphs_info = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        tables_info = []
        for idx, tbl in enumerate(doc.tables, 1):
            if len(tbl.rows) > 0:
                headers = [c.text.strip().replace("\n", " ") for c in tbl.rows[0].cells]
                sample_row = [c.text.strip().replace("\n", " ") for c in tbl.rows[1].cells] if len(tbl.rows) > 1 else []
                tables_info.append({
                    "bang_so": idx,
                    "so_cot": len(tbl.columns),
                    "tieu_de_cot": headers,
                    "dong_mau_tham_khao": sample_row
                })
        summary = {
            "cac_muc_van_ban_trong_form": paragraphs_info[:12],
            "cau_truc_cac_bang_trong_form": tables_info
        }
        return json.dumps(summary, ensure_ascii=False, indent=2)
    except Exception as e:
        return f"Không đọc được chi tiết cấu trúc mẫu: {e}"

# ================= HÀM QUÉT MODEL & TỰ ĐỘNG FALLBACK KHI LỖI 429 =================
@st.cache_data(ttl=300, show_spinner=False)
def validate_and_fetch_models(raw_key):
    api_key = clean_api_key(raw_key)
    if not api_key:
        return False, [], "Chưa nhập API Key."
    try:
        genai.configure(api_key=api_key)
        models = []
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods:
                models.append(m.name.replace("models/", ""))
        if not models:
            return False, [], "API Key không trả về model nào."
        return True, models, f"API Key chuẩn ({len(models)} model khả dụng)"
    except Exception as e:
        return False, [], f"API Key không hợp lệ: {e}"

def pick_optimal_model(models_list, task="recap"):
    if not models_list:
        return "gemini-2.5-flash"
    priorities = [
        "gemini-2.5-flash",
        "gemini-3.5-flash",
        "gemini-3.1-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-3.1-pro-preview"
    ]
    for p in priorities:
        for m in models_list:
            if m == p or m.startswith(p):
                return m
    return models_list[0]

def upload_large_audio_rest(api_key, file_bytes, mime_type, display_name="meeting_audio"):
    init_url = f"https://generativelanguage.googleapis.com/upload/v1beta/files?key={api_key}"
    headers = {
        "X-Goog-Upload-Protocol": "resumable",
        "X-Goog-Upload-Command": "start",
        "X-Goog-Upload-Header-Content-Length": str(len(file_bytes)),
        "X-Goog-Upload-Header-Content-Type": mime_type,
        "Content-Type": "application/json"
    }
    r_init = requests.post(init_url, headers=headers, json={"file": {"display_name": display_name}}, timeout=30)
    r_init.raise_for_status()
    upload_url = r_init.headers.get("X-Goog-Upload-URL")
    if not upload_url:
        raise RuntimeError("Không lấy được đường dẫn Resumable Upload từ Google.")

    upload_headers = {
        "Content-Length": str(len(file_bytes)),
        "X-Goog-Upload-Offset": "0",
        "X-Goog-Upload-Command": "upload, finalize"
    }
    r_up = requests.post(upload_url, headers=upload_headers, data=file_bytes, timeout=300)
    r_up.raise_for_status()
    file_info = r_up.json().get("file", {})
    file_uri = file_info.get("uri")
    file_name = file_info.get("name")

    for _ in range(60):
        check_res = requests.get(f"https://generativelanguage.googleapis.com/v1beta/{file_name}?key={api_key}", timeout=15)
        if check_res.status_code == 200:
            state = check_res.json().get("state")
            if state == "ACTIVE":
                break
            elif state == "FAILED":
                raise RuntimeError("Google báo lỗi khi giải mã file âm thanh này.")
        time.sleep(2)
    return file_uri, file_name

def generate_recap_with_audio_or_text(api_key, model_name, prompt_text, audio_file=None):
    genai.configure(api_key=api_key)
    fallback_chain = [
        model_name,
        "gemini-2.5-flash",
        "gemini-3.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash"
    ]
    candidate_models = list(dict.fromkeys(fallback_chain))

    if audio_file is None:
        last_err = None
        for m_name in candidate_models:
            try:
                model = genai.GenerativeModel(m_name)
                res = model.generate_content(prompt_text)
                return res.text
            except Exception as e:
                last_err = e
                if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e) or "404" in str(e):
                    continue
                raise e
        raise RuntimeError(f"Tất cả model đều vượt hạn mức miễn phí: {last_err}")

    file_bytes = audio_file.getvalue()
    ext = audio_file.name.split(".")[-1].lower()
    mime_map = {"webm": "audio/webm", "wav": "audio/wav", "mp3": "audio/mp3", "m4a": "audio/mp4", "ogg": "audio/ogg"}
    mime_type = mime_map.get(ext, "audio/mp4")
    size_mb = len(file_bytes) / (1024 * 1024)

    if size_mb <= 18.0:
        audio_part = {"mime_type": mime_type, "data": file_bytes}
        last_err = None
        for m_name in candidate_models:
            try:
                model = genai.GenerativeModel(m_name)
                res = model.generate_content([prompt_text, audio_part])
                return res.text
            except Exception as e:
                last_err = e
                if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e) or "404" in str(e):
                    continue
                raise e
        raise RuntimeError(f"Tất cả model đều vượt hạn mức miễn phí: {last_err}")

    file_uri, file_name = upload_large_audio_rest(api_key, file_bytes, mime_type, audio_file.name)
    try:
        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt_text},
                    {"file_data": {"mime_type": mime_type, "file_uri": file_uri}}
                ]
            }]
        }
        last_error_text = ""
        for m_name in candidate_models:
            gen_url = f"https://generativelanguage.googleapis.com/v1beta/models/{m_name}:generateContent?key={api_key}"
            r_gen = requests.post(gen_url, json=payload, timeout=600)
            if r_gen.status_code == 200:
                data = r_gen.json()
                return data["candidates"][0]["content"]["parts"][0]["text"]
            last_error_text = f"Lỗi {r_gen.status_code} ({m_name}): {r_gen.text}"
            if r_gen.status_code in [429, 404, 503]:
                continue
            else:
                raise RuntimeError(last_error_text)
        raise RuntimeError(f"Đã thử tự động chuyển đổi qua các model nhưng đều chạm hạn mức: {last_error_text}")
    finally:
        try:
            requests.delete(f"https://generativelanguage.googleapis.com/v1beta/{file_name}?key={api_key}", timeout=10)
        except Exception:
            pass

# ================= HÀM DỊCH THUẬT AI + SO KHỚP GLOSSARY, TM & NHÂN SỰ =================
def translate_batch_with_thaco_memory(lines_with_time, mode, raw_key, model_name, raw_url=""):
    api_key = clean_api_key(raw_key)
    results = []
    tm_to_upsert = {}

    for time_tag, clean_text in lines_with_time:
        if not clean_text:
            continue
        if mode == "Việt -> Việt (Chỉ ghi phụ đề)":
            results.append({"time": time_tag, "source": clean_text, "target": clean_text, "terms": []})
            continue

        if clean_text in st.session_state.tm:
            results.append({"time": time_tag, "source": clean_text, "target": st.session_state.tm[clean_text], "terms": ["⚡ TM Dùng Chung"]})
            continue

        matched = {}
        for vn_term, en_term in st.session_state.glossary.items():
            if vn_term.lower() in clean_text.lower() or en_term.lower() in clean_text.lower():
                matched[vn_term] = en_term

        if not api_key:
            results.append({"time": time_tag, "source": clean_text, "target": "[Chưa nhập Gemini API Key]", "terms": list(matched.keys())})
            continue

        direction = "tiếng Anh sang tiếng Việt" if mode == "Anh -> Việt" else "tiếng Việt sang tiếng Anh"
        glossary_instruction = f"BẮT BUỘC dùng đúng thuật ngữ THACO sau: {json.dumps(matched, ensure_ascii=False)}. " if matched else ""
        personnel_context = f"Danh sách nhân sự/phòng ban THACO để tham chiếu: {json.dumps(st.session_state.personnel, ensure_ascii=False)}. "

        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            prompt = (
                f"Bạn là chuyên gia phiên dịch kỹ thuật và quản lý chất lượng của THACO AUTO. "
                f"Hãy dịch câu hội thoại sau từ {direction} thật chuẩn xác, ngắn gọn, đúng văn phong kỹ thuật ô tô. "
                f"{glossary_instruction}{personnel_context}"
                f"Chỉ trả về duy nhất câu bản dịch, không giải thích:\n{clean_text}"
            )
            res = model.generate_content(prompt)
            translated = res.text.strip()
            results.append({"time": time_tag, "source": clean_text, "target": translated, "terms": list(matched.keys())})
            if len(clean_text.split()) >= 3:
                tm_to_upsert[clean_text] = translated
        except Exception as e:
            results.append({"time": time_tag, "source": clean_text, "target": f"[Lỗi AI: {e}]", "terms": []})

    if raw_url and tm_to_upsert:
        upsert_tm_to_sheets(raw_url, tm_to_upsert)
    return results

# ================= HÀM ĐIỀN BẢNG THÀNH PHẦN THAM DỰ & NỘI DUNG VÀO WORD BBH =================
def set_cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tcBorders.append(border)
    tcPr.append(tcBorders)

def build_attendance_matrix(selected_personnel):
    grouped = {}
    for p in selected_personnel:
        dept = p.get("phong_ban", "").strip() or "Đơn vị tham dự"
        grouped.setdefault(dept, []).append(f"{p.get('ho_ten', '')} – {p.get('chuc_danh', '')}".strip(" –"))
    if not grouped:
        return [], []
    depts = list(grouped.keys())
    max_rows = max(len(v) for v in grouped.values())
    matrix = []
    for r in range(max_rows):
        row_vals = []
        for d in depts:
            row_vals.append(grouped[d][r] if r < len(grouped[d]) else "")
        matrix.append(row_vals)
    return depts, matrix

def populate_uploaded_thaco_docx(template_file_bytes, meta, recap_rows, selected_personnel, update_attendance_table=True):
    doc = Document(io.BytesIO(template_file_bytes))

    # 1. Cập nhật Số BBH ở bảng Header
    if len(doc.tables) >= 1:
        header_tbl = doc.tables[0]
        for row in header_tbl.rows:
            for cell in row.cells:
                if "Số:" in cell.text or "BBH" in cell.text:
                    for p in cell.paragraphs:
                        if "Số:" in p.text or "BBH" in p.text:
                            p.text = f"Số: {meta['so_bbh']}"
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 2. Cập nhật V/v, Thời gian, Địa điểm, Kết luận
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.startswith("V/v:") or txt.startswith("V/v :"):
            p.text = f"V/v: {meta['ve_viec']}"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
        elif txt.startswith("Thời gian:"):
            p.text = f"Thời gian: {meta['thoi_gian']}"
        elif txt.startswith("Địa điểm:"):
            p.text = f"Địa điểm: {meta['dia_diem']}"
        elif txt.startswith("Trên đây là các thống nhất"):
            p.text = meta["ket_luan_chung"]

    # 3. Cập nhật Bảng II. Thành phần tham dự (Bảng thứ 2 - index 1)
    if update_attendance_table and selected_personnel and len(doc.tables) >= 2:
        att_tbl = doc.tables[1]
        depts, matrix = build_attendance_matrix(selected_personnel)
        if depts:
            while len(att_tbl.rows) > 0:
                att_tbl._tbl.remove(att_tbl.rows[0]._tr)
            num_cols = len(att_tbl.columns)
            hdr_cells = att_tbl.add_row().cells
            for c_i in range(num_cols):
                hdr_cells[c_i].text = depts[c_i] if c_i < len(depts) else ""
                for run in hdr_cells[c_i].paragraphs[0].runs:
                    run.bold = True
                set_cell_border(hdr_cells[c_i])
            for r_vals in matrix:
                r_cells = att_tbl.add_row().cells
                for c_i in range(num_cols):
                    r_cells[c_i].text = r_vals[c_i] if c_i < len(r_vals) else ""
                    set_cell_border(r_cells[c_i])

    # 4. Tìm và điền Bảng III: Nội dung thảo luận (7 cột)
    target_table = None
    for tbl in doc.tables:
        if len(tbl.columns) >= 6 and len(tbl.rows) >= 1:
            header_text = " ".join([c.text.strip().lower() for c in tbl.rows[0].cells])
            if "vấn đề" in header_text or "thảo luận" in header_text or "tt" in header_text:
                target_table = tbl
                break
    if target_table is None and len(doc.tables) >= 3:
        for tbl in doc.tables:
            if len(tbl.columns) == 7:
                target_table = tbl
                break

    if target_table is not None:
        while len(target_table.rows) > 1:
            target_table._tbl.remove(target_table.rows[1]._tr)

        for idx, item in enumerate(recap_rows, 1):
            row_cells = target_table.add_row().cells
            values = [
                str(item.get("tt") or idx),
                str(item.get("van_de", "")),
                str(item.get("noi_dung_thao_luan", "")),
                str(item.get("cong_viec_tiep_theo", "")),
                str(item.get("trach_nhiem", "")),
                str(item.get("phoi_hop", "-")),
                str(item.get("hoan_thanh", ""))
            ]
            for c_idx in range(min(len(row_cells), 7)):
                row_cells[c_idx].text = values[c_idx]
                set_cell_border(row_cells[c_idx])
                if c_idx in [0, 4, 5, 6]:
                    for p in row_cells[c_idx].paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    out_bio = io.BytesIO()
    doc.save(out_bio)
    out_bio.seek(0)
    return out_bio.getvalue()

def generate_default_thaco_bbh_docx(meta, recap_rows, selected_personnel):
    doc = Document()
    hdr_table = doc.add_table(rows=1, cols=2)
    hdr_table.cell(0, 0).text = "CÔNG TY TNHH THACO AUTO"
    hdr_table.cell(0, 1).text = f"Số: {meta['so_bbh']}"
    
    title = doc.add_heading("BIÊN BẢN HỌP", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_vv = doc.add_paragraph(f"V/v: {meta['ve_viec']}")
    p_vv.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("I. Thời gian, địa điểm:", level=2)
    doc.add_paragraph(f"Thời gian: {meta['thoi_gian']}\nĐịa điểm: {meta['dia_diem']}")

    doc.add_heading("II. Thành phần tham dự:", level=2)
    depts, matrix = build_attendance_matrix(selected_personnel)
    if depts:
        att_tbl = doc.add_table(rows=1, cols=len(depts))
        for c_i, d_name in enumerate(depts):
            att_tbl.rows[0].cells[c_i].text = d_name
            set_cell_border(att_tbl.rows[0].cells[c_i])
        for r_vals in matrix:
            row_cells = att_tbl.add_row().cells
            for c_i, val in enumerate(r_vals):
                row_cells[c_i].text = val
                set_cell_border(row_cells[c_i])

    doc.add_heading("III. Nội dung thảo luận và kế hoạch triển khai thực hiện:", level=2)
    main_tbl = doc.add_table(rows=1, cols=7)
    headers = ["TT", "Vấn đề", "Nội dung thảo luận", "Công việc thực hiện tiếp theo", "T.nhiệm", "P.hợp", "H.thành"]
    for i, h in enumerate(headers):
        main_tbl.rows[0].cells[i].text = h
        set_cell_border(main_tbl.rows[0].cells[i])

    for idx, item in enumerate(recap_rows, 1):
        row_cells = main_tbl.add_row().cells
        vals = [
            str(item.get("tt") or idx),
            str(item.get("van_de", "")),
            str(item.get("noi_dung_thao_luan", "")),
            str(item.get("cong_viec_tiep_theo", "")),
            str(item.get("trach_nhiem", "")),
            str(item.get("phoi_hop", "-")),
            str(item.get("hoan_thanh", ""))
        ]
        for i in range(7):
            row_cells[i].text = vals[i]
            set_cell_border(row_cells[i])

    doc.add_paragraph("\n" + meta["ket_luan_chung"])

    sign_tbl = doc.add_table(rows=2, cols=4)
    sign_headers = ["Khối KD Xuất khẩu", "Nhà máy Bus", "Nhà máy Tải", "P.QLCL KCN"]
    for i, sh in enumerate(sign_headers):
        sign_tbl.cell(0, i).text = sh
        sign_tbl.cell(1, i).text = "\n\n\n"

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# ================= THANH BÊN (SIDEBAR) =================
with st.sidebar:
    st.header("⚙️ Cấu Hình Hệ Thống")
    
    default_key = st.secrets.get("GEMINI_API_KEY", "") if "GEMINI_API_KEY" in st.secrets else ""
    default_url = st.secrets.get("WEBAPP_URL", "") if "WEBAPP_URL" in st.secrets else ""

    raw_api_key_input = st.text_input("🔑 1. Gemini API Key:", value=default_key, type="password")
    gemini_api_key = clean_api_key(raw_api_key_input)
    
    if gemini_api_key:
        is_valid, models_list, key_msg = validate_and_fetch_models(gemini_api_key)
        if is_valid:
            st.success(f"✅ {key_msg}")
            opt_recap = pick_optimal_model(models_list, task="recap")
            opt_trans = pick_optimal_model(models_list, task="translate")
            
            selected_recap_model = st.selectbox(
                "🧠 Model AI cho Recap & Nghe File Âm Thanh:",
                options=models_list,
                index=models_list.index(opt_recap) if opt_recap in models_list else 0
            )
            selected_trans_model = st.selectbox(
                "⚡ Model AI cho Dịch Phụ Đề Nhanh:",
                options=models_list,
                index=models_list.index(opt_trans) if opt_trans in models_list else 0
            )
        else:
            st.error(f"❌ {key_msg}")
            selected_recap_model = "gemini-2.5-flash"
            selected_trans_model = "gemini-2.5-flash"
    else:
        st.caption("Nhập Gemini API Key để hiển thị danh sách Model.")
        selected_recap_model = "gemini-2.5-flash"
        selected_trans_model = "gemini-2.5-flash"

    st.divider()
    raw_webapp_input = st.text_input("☁️ 2. Link Google Apps Script (Dùng chung):", value=default_url)
    webapp_url = clean_webapp_url(raw_webapp_input)
    
    if webapp_url and not st.session_state.is_synced:
        ok, msg = sync_from_sheets(webapp_url)
        if ok:
            st.toast(msg)

    if st.button("🔄 Đồng bộ Glossary, TM & Nhân sự từ Sheets", use_container_width=True, type="primary"):
        ok, msg = sync_from_sheets(webapp_url)
        if ok:
            st.success(msg)
        else:
            st.error(msg)

    st.divider()
    st.subheader("🌐 Chế độ Ngôn ngữ Họp")
    trans_mode = st.radio(
        "Chọn chiều nhận diện & dịch:",
        ["Anh -> Việt", "Việt -> Anh", "Việt -> Việt (Chỉ ghi phụ đề)"]
    )
    
    st.divider()
    with st.expander("👤 Thêm nhanh Nhân sự vào Google Sheets"):
        p_name = st.text_input("Họ và tên (VD: Ông Võ Thành Lưu):")
        p_role = st.text_input("Chức danh (VD: TP QLCL xe Thương mại):")
        p_dept = st.text_input("Phòng ban / Đơn vị (VD: Phòng QLCL KCN):")
        if st.button("💾 Lưu Nhân sự lên Sheets", use_container_width=True):
            if p_name and webapp_url:
                updated_p = st.session_state.personnel + [{"ho_ten": p_name.strip(), "chuc_danh": p_role.strip(), "phong_ban": p_dept.strip()}]
                ok, msg = save_personnel_to_sheets(webapp_url, updated_p)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
            else:
                st.warning("Vui lòng nhập Họ tên và Link Web App URL.")

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

# ================= THANH HEADER EXECUTIVE CHUẨN THACO AUTO =================
st.markdown(f"""
<div class="thaco-header">
    <div class="thaco-title">🚗 HỆ THỐNG TRỢ LÝ HỌP, DỊCH THUẬT REAL-TIME & XUẤT BBH — LƯU PRO</div>
    <div class="thaco-badges">
        <span class="thaco-badge">🤖 Recap: {selected_recap_model}</span>
        <span class="thaco-badge">📖 Glossary: {len(st.session_state.glossary)}</span>
        <span class="thaco-badge">🧠 TM: {len(st.session_state.tm)}</span>
        <span class="thaco-badge">👥 Nhân sự: {len(st.session_state.personnel)}</span>
    </div>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "🎙️ Bước 1: Cockpit Thu Âm & Dịch Real-time",
    "🧠 Bước 2: Thiết Lập Nhân Sự, Mẫu BBH & AI Recap",
    "📄 Bước 3: Xuất Biên Bản Word Trình Ký",
    "📚 Quản Trị Nhân Sự, TM & Glossary (Google Sheets)"
])

# ----------------- TAB 1: COCKPIT THU ÂM VB-CABLE & BÀN DỊCH SONG SONG -----------------
with tab1:
    col_left_cockpit, col_right_cockpit = st.columns([1.25, 1], gap="medium")

    with col_left_cockpit:
        speech_lang = "en-US" if trans_mode == "Anh -> Việt" else "vi-VN"
        glossary_json_str = json.dumps(st.session_state.glossary, ensure_ascii=False)
        tm_json_str = json.dumps(st.session_state.tm, ensure_ascii=False)
        personnel_json_str = json.dumps(st.session_state.personnel, ensure_ascii=False)

        web_audio_component = f"""
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-radius:12px; padding:14px; font-family:system-ui,-apple-system,sans-serif; box-shadow:0 2px 8px rgba(0,0,0,0.04);">
          <div style="display:flex; gap:8px; flex-wrap:wrap; align-items:center; margin-bottom:10px;">
            <select id="audioSource" style="flex:1; min-width:190px; padding:7px 10px; border-radius:6px; border:1px solid #94a3b8; font-weight:600; font-size:13px;"></select>
            <button id="btnStart" onclick="startCapture()" style="background:#16a34a; color:white; border:none; padding:8px 14px; border-radius:6px; font-weight:700; cursor:pointer; font-size:13px;">▶️ Thu & Lọc Nhiễu</button>
            <button id="btnStop" onclick="stopCapture()" disabled style="background:#dc2626; color:white; border:none; padding:8px 14px; border-radius:6px; font-weight:700; cursor:pointer; opacity:0.4; font-size:13px;">⏹️ Dừng</button>
            <button onclick="openFloatingPip()" title="Tách phụ đề thành cửa sổ nổi luôn nằm trên MS Teams/Zoom" style="background:#003366; color:white; border:none; padding:8px 12px; border-radius:6px; font-weight:600; cursor:pointer; font-size:13px;">🪟 Phụ đề Nổi (Teams/Zoom)</button>
            <span id="timerDisplay" style="background:#0f172a; color:#38bdf8; padding:6px 10px; border-radius:6px; font-family:monospace; font-weight:bold; font-size:14px;">00:00:00</span>
          </div>

          <div style="margin-bottom:10px;">
            <div style="font-size:12px; color:#475569; font-weight:600; margin-bottom:5px;">🗣️ Bấm chọn nhanh Người đang phát biểu (1-Click Tag):</div>
            <div id="speakerPills" style="display:flex; gap:6px; flex-wrap:wrap;"></div>
          </div>

          <div style="display:flex; align-items:center; gap:10px; margin-bottom:8px; font-size:12px; color:#334155;">
            <span>🎚️ <b>DSP VB-Cable:</b></span>
            <div style="flex:1; height:8px; background:#e2e8f0; border-radius:4px; overflow:hidden;">
              <div id="volumeBar" style="width:0%; height:100%; background:linear-gradient(90deg,#22c55e,#eab308); transition:width 0.08s;"></div>
            </div>
            <span id="statusText" style="font-weight:bold; color:#64748b;">Chờ kích hoạt</span>
            <a id="downloadAudio" style="display:none; background:#2563eb; color:white; text-decoration:none; padding:4px 10px; border-radius:5px; font-weight:bold;">💾 Audio (.webm)</a>
            <a id="downloadTxt" style="display:none; background:#0d9488; color:white; text-decoration:none; padding:4px 10px; border-radius:5px; font-weight:bold;">📄 Phụ đề (.txt)</a>
          </div>

          <div id="liveSubBox" style="height:310px; overflow-y:auto; background:#0f172a; color:#f8fafc; padding:14px; border-radius:8px; font-size:14.5px; line-height:1.65; border:1px solid #1e293b;">
            <div id="finalSubtitles"></div>
            <div id="interimSubtitle" style="color:#38bdf8; font-style:italic;"></div>
          </div>

          <div style="margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
            <span style="font-size:12px; color:#64748b;">💡 Bấm <b>Copy sang Bàn Dịch</b> để đẩy toàn bộ sang cột bên phải:</span>
            <button onclick="copyTranscript()" style="background:#0284c7; color:white; border:none; padding:6px 14px; border-radius:6px; font-weight:bold; cursor:pointer; font-size:12.5px;">📋 Copy toàn bộ sang Bàn Dịch ➔</button>
          </div>
          <textarea id="fullBuffer" style="display:none;"></textarea>
        </div>

        <script>
          let mediaStream = null, mediaRecorder = null, recordedChunks = [], recognition = null;
          let isCapturing = false, audioContext = null, analyser = null;
          let currentSpeaker = "", timerInterval = null, secondsElapsed = 0, pipWindow = null;
          const glossary = {glossary_json_str};
          const tm = {tm_json_str};
          const personnelList = {personnel_json_str};

          function renderSpeakerPills() {{
            const container = document.getElementById('speakerPills');
            container.innerHTML = '';
            const allOptions = [{{ label: 'Tự động / Chung', val: '' }}].concat(
              personnelList.map(p => ({{
                label: p.ho_ten.replace(/^(Ông|Bà|Anh|Chị)\\s+/i, '') + ' (' + p.phong_ban + ')',
                val: `${{p.ho_ten}} – ${{p.chuc_danh}} (${{p.phong_ban}})`
              }}))
            );
            allOptions.forEach(item => {{
              const btn = document.createElement('button');
              btn.innerText = item.val === '' ? '🔄 Tự động' : '👤 ' + item.label;
              const active = currentSpeaker === item.val;
              btn.style.cssText = `padding:5px 10px; border-radius:16px; font-size:12px; font-weight:600; cursor:pointer; border:1px solid ${{active ? '#0284c7' : '#cbd5e1'}}; background:${{active ? '#0284c7' : '#f8fafc'}}; color:${{active ? '#fff' : '#334155'}}; transition:0.15s;`;
              btn.onclick = () => {{ currentSpeaker = item.val; renderSpeakerPills(); }};
              container.appendChild(btn);
            }});
          }}
          renderSpeakerPills();

          function formatTime(sec) {{
            const h = String(Math.floor(sec / 3600)).padStart(2, '0');
            const m = String(Math.floor((sec % 3600) / 60)).padStart(2, '0');
            const s = String(sec % 60).padStart(2, '0');
            return `${{h}}:${{m}}:${{s}}`;
          }}

          async function openFloatingPip() {{
            if ('documentPictureInPicture' in window) {{
              pipWindow = await window.documentPictureInPicture.requestWindow({{ width: 650, height: 240 }});
              pipWindow.document.body.style.cssText = "background:#0f172a; color:#f8fafc; font-family:sans-serif; padding:12px; margin:0; overflow-y:auto;";
              pipWindow.document.body.innerHTML = `<div style="font-size:12px; color:#38bdf8; font-weight:bold; margin-bottom:6px;">🚗 THACO LIVE SUBTITLES (LUÔN NỔI TRÊN TEAMS/ZOOM)</div><div id="pipContent" style="font-size:16px; line-height:1.5;"></div>`;
            }} else {{
              alert("Trình duyệt cần là Chrome/Edge bản mới để bật Cửa sổ nổi Picture-in-Picture.");
            }}
          }}

          async function loadAudioDevices() {{
            try {{
              await navigator.mediaDevices.getUserMedia({{ audio: true }});
              const devices = await navigator.mediaDevices.enumerateDevices();
              const select = document.getElementById('audioSource');
              select.innerHTML = '';
              devices.filter(d => d.kind === 'audioinput').forEach(d => {{
                const opt = document.createElement('option');
                opt.value = d.deviceId;
                opt.text = d.label || 'Audio Input';
                if (opt.text.toLowerCase().includes('cable')) {{
                  opt.selected = true;
                  opt.text = '⭐ ' + opt.text;
                }}
                select.appendChild(opt);
              }});
            }} catch (e) {{}}
          }}
          loadAudioDevices();

          function enrichText(text) {{
            let html = text;
            personnelList.forEach(p => {{
              const cleanName = p.ho_ten.replace(/^(Ông|Bà|Anh|Chị)\\s+/i, '').trim();
              if (cleanName.length > 3 && text.toLowerCase().includes(cleanName.toLowerCase())) {{
                html += ` <span style="background:#059669; color:#fff; padding:1px 6px; border-radius:4px; font-size:11.5px;">[👤 ${{p.ho_ten}} - ${{p.phong_ban}}]</span>`;
              }}
            }});
            if (tm[text]) {{
              html += `<br><span style="color:#4ade80; font-weight:bold;">⚡ [TM]: ${{tm[text]}}</span>`;
            }}
            for (const [vn, en] of Object.entries(glossary)) {{
              if (vn && text.toLowerCase().includes(vn.toLowerCase())) {{
                html += ` <span style="background:#ca8a04; color:#fff; padding:1px 6px; border-radius:4px; font-size:11.5px;">[${{vn}} ➔ ${{en}}]</span>`;
              }} else if (en && text.toLowerCase().includes(en.toLowerCase())) {{
                html += ` <span style="background:#ca8a04; color:#fff; padding:1px 6px; border-radius:4px; font-size:11.5px;">[${{en}} ➔ ${{vn}}]</span>`;
              }}
            }}
            return html;
          }}

          function updateTxtDownloadLink() {{
            const textVal = document.getElementById('fullBuffer').value;
            if (!textVal.trim()) return;
            const blobTxt = new Blob([textVal], {{ type: 'text/plain;charset=utf-8' }});
            const aTxt = document.getElementById('downloadTxt');
            aTxt.href = URL.createObjectURL(blobTxt);
            aTxt.download = 'Phu_De_Hop_THACO_' + new Date().toISOString().slice(0,19).replace(/:/g,'-') + '.txt';
            aTxt.style.display = 'inline-block';
          }}

          async function startCapture() {{
            const deviceId = document.getElementById('audioSource').value;
            try {{
              mediaStream = await navigator.mediaDevices.getUserMedia({{
                audio: {{ deviceId: deviceId ? {{ exact: deviceId }} : undefined, noiseSuppression: true, echoCancellation: true, autoGainControl: true }}
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
                let avg = Math.min(100, Math.round((dataArray.reduce((a, b) => a + b, 0) / dataArray.length) * 1.8));
                document.getElementById('volumeBar').style.width = avg + '%';
                requestAnimationFrame(updateVolume);
              }}

              recordedChunks = [];
              mediaRecorder = new MediaRecorder(mediaStream);
              mediaRecorder.ondataavailable = e => {{ if (e.data.size > 0) recordedChunks.push(e.data); }};
              mediaRecorder.onstop = () => {{
                const blob = new Blob(recordedChunks, {{ type: 'audio/webm' }});
                const a = document.getElementById('downloadAudio');
                a.href = URL.createObjectURL(blob);
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
                const spkBadge = currentSpeaker ? `<span style="color:#38bdf8; font-weight:bold;">[${{currentSpeaker}}]</span> ` : '';
                const spkTxt = currentSpeaker ? `[${{currentSpeaker}}] ` : '';

                for (let i = event.resultIndex; i < event.results.length; ++i) {{
                  const transcript = event.results[i][0].transcript.trim();
                  if (event.results[i].isFinal) {{
                    const now = new Date().toLocaleTimeString('vi-VN');
                    const decorated = enrichText(transcript);
                    const rowHtml = `<div style="margin-bottom:7px; border-bottom:1px dashed #334155; padding-bottom:4px;"><span style="color:#94a3b8;">[${{now}}]</span> ${{spkBadge}}${{decorated}}</div>`;
                    const finalDiv = document.getElementById('finalSubtitles');
                    finalDiv.innerHTML += rowHtml;
                    const box = document.getElementById('liveSubBox');
                    box.scrollTop = box.scrollHeight;

                    if (pipWindow && !pipWindow.closed) {{
                      const pipDiv = pipWindow.document.getElementById('pipContent');
                      if (pipDiv) {{ pipDiv.innerHTML = rowHtml; }}
                    }}

                    const buf = document.getElementById('fullBuffer');
                    buf.value += `[${{now}}] ${{spkTxt}}${{transcript}}\\n`;
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

              secondsElapsed = 0;
              clearInterval(timerInterval);
              timerInterval = setInterval(() => {{
                secondsElapsed++;
                document.getElementById('timerDisplay').innerText = formatTime(secondsElapsed);
              }}, 1000);

              document.getElementById('btnStart').disabled = true;
              document.getElementById('btnStart').style.opacity = '0.4';
              document.getElementById('btnStop').disabled = false;
              document.getElementById('btnStop').style.opacity = '1';
              document.getElementById('statusText').innerText = '🟢 Đang thu';
              document.getElementById('statusText').style.color = '#16a34a';
            }} catch (err) {{
              alert('Không thể mở thiết bị âm thanh: ' + err.message);
            }}
          }}

          function stopCapture() {{
            isCapturing = false;
            clearInterval(timerInterval);
            if (recognition) recognition.stop();
            if (mediaRecorder && mediaRecorder.state !== 'inactive') mediaRecorder.stop();
            if (mediaStream) mediaStream.getTracks().forEach(t => t.stop());
            if (audioContext) audioContext.close();
            updateTxtDownloadLink();
            document.getElementById('btnStart').disabled = false;
            document.getElementById('btnStart').style.opacity = '1';
            document.getElementById('btnStop').disabled = true;
            document.getElementById('btnStop').style.opacity = '0.4';
            document.getElementById('statusText').innerText = '⏹️ Đã dừng';
            document.getElementById('statusText').style.color = '#dc2626';
          }}

          function copyTranscript() {{
            const textVal = document.getElementById('fullBuffer').value;
            navigator.clipboard.writeText(textVal);
            alert('Đã copy! Hãy bấm Ctrl+V vào ô bên phải để dịch & chuẩn hóa.');
          }}
        </script>
        """
        components.html(web_audio_component, height=525)

    with col_right_cockpit:
        st.markdown("##### 🔄 Bàn Dịch Thuật & Hiệu Đính Đồng Bộ (`upsert_tm`)")
        raw_input_box = st.text_area(
            "📥 Dán nội dung vừa Copy (Ctrl+V) hoặc gõ câu cần dịch:",
            value=st.session_state.raw_stream_input,
            height=135,
            placeholder="[09:15:20] [Ông Võ Thành Lưu – TP QLCL xe Thương mại] Đề nghị BJM gửi báo cáo kỹ thuật..."
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

            with st.spinner(f"Đang dịch bằng {selected_trans_model}..."):
                st.session_state.transcript_list = translate_batch_with_thaco_memory(
                    parsed_lines, trans_mode, gemini_api_key, selected_trans_model, webapp_url
                )

        if st.session_state.transcript_list:
            edited_subtitles = st.data_editor(
                st.session_state.transcript_list,
                column_config={
                    "time": st.column_config.TextColumn("Giờ", width="small", disabled=True),
                    "source": st.column_config.TextColumn("Câu gốc", width="medium"),
                    "target": st.column_config.TextColumn("Bản Dịch Chuẩn (Sửa trực tiếp)", width="large"),
                    "terms": st.column_config.ListColumn("Thuật ngữ", width="small")
                },
                use_container_width=True,
                height=245,
                key="sub_editor"
            )
            st.session_state.transcript_list = edited_subtitles

            if st.button("💾 Lưu bản dịch vừa hiệu đính đè lên Google Sheets (`upsert_tm`)", use_container_width=True):
                upsert_dict = {item["source"]: item["target"] for item in edited_subtitles if item.get("source") and item.get("target")}
                ok, msg = upsert_tm_to_sheets(webapp_url, upsert_dict)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
        else:
            st.info("💡 Khung Hiệu đính trực tiếp sẽ hiển thị ngay tại đây sau khi anh bấm nút Dịch ở trên.")

# ----------------- TAB 2: THIẾT LẬP NHÂN SỰ, MẪU BBH ĐẦU VÀO & AI RECAP -----------------
with tab2:
    st.markdown("""
    <div class="step-card">
        <b>🎯 Quy trình AI Recap Chuẩn Xác Theo Form & Nhân Sự:</b><br>
        1️⃣ <b>Chọn người tham gia & Tải lên Mẫu BBH (.docx):</b> Giúp AI biết chính xác những ai có mặt trong cuộc họp và cấu trúc các cột/bảng của Mẫu BBH.<br>
        2️⃣ <b>Chọn nguồn dữ liệu họp (File Ghi Âm / .TXT):</b> AI sẽ đối chiếu giọng nói/nội dung với đúng tên Nhân sự, Phòng ban và Bảng thuật ngữ THACO để điền vào Form.
    </div>
    """, unsafe_allow_html=True)

    # KHỐI 1: CHỌN NHÂN SỰ THAM GIA & TẢI MẪU BBH ĐẦU VÀO TRƯỚC KHI RECAP
    col_setup_p, col_setup_tpl = st.columns([1.2, 1], gap="medium")

    with col_setup_p:
        st.markdown("##### 👥 1. Chọn Nhân sự tham gia cuộc họp (Đầu vào cho AI & Bảng Thành phần)")
        all_p_labels = [
            f"{p.get('ho_ten', '')} – {p.get('chuc_danh', '')} ({p.get('phong_ban', '')})"
            for p in st.session_state.personnel
        ]
        selected_p_labels_step2 = st.multiselect(
            "Tick chọn những nhân sự tham dự cuộc họp hôm nay:",
            options=all_p_labels,
            default=all_p_labels,
            key="multiselect_attendees_step2"
        )
        st.session_state.selected_attendees = [
            p for p, lbl in zip(st.session_state.personnel, all_p_labels) if lbl in selected_p_labels_step2
        ]
        st.session_state.extra_attendees_str = st.text_input(
            "➕ Nhân sự / Đối tác bổ sung ngoài danh bạ (nếu có, phân cách bằng dấu phẩy):",
            value=st.session_state.extra_attendees_str,
            placeholder="VD: Mr. Wang - Kỹ sư BJM, Ông Lê Văn A - NM Bus..."
        )

    with col_setup_tpl:
        st.markdown("##### 📄 2. Tải lên Mẫu BBH (.docx) làm chuẩn đầu vào cho AI")
        uploaded_tpl_step2 = st.file_uploader(
            "Tải file Word BBH Mẫu (VD: `2026-1113. BB họp chất lượng xe xuất khẩuV4 - Copy.docx`):",
            type=["docx"],
            key="uploader_tpl_step2"
        )
        if uploaded_tpl_step2 is not None:
            st.session_state.template_docx_bytes = uploaded_tpl_step2.getvalue()
            st.session_state.template_docx_name = uploaded_tpl_step2.name
            st.session_state.template_structure_hint = inspect_docx_template(st.session_state.template_docx_bytes)
            st.success(f"✅ Đã nạp mẫu `{uploaded_tpl_step2.name}`! AI sẽ bám sát cấu trúc mẫu này và dùng luôn cho Bước 3.")
        elif st.session_state.template_docx_bytes is not None:
            st.info(f"📌 Đang sử dụng mẫu đã nạp: `{st.session_state.template_docx_name}`")
        else:
            st.caption("Nếu chưa tải file mẫu, hệ thống sẽ dùng cấu trúc chuẩn 7 cột của Phòng QLCL KCN - THACO AUTO.")

    st.divider()

    # KHỐI 2: CHỌN NGUỒN DỮ LIỆU CUỘC HỌP & PROMPT RECAP
    col_src, col_prm = st.columns([1, 1.2], gap="medium")

    with col_src:
        st.markdown("##### 🎙️ 3. Chọn nguồn dữ liệu cuộc họp")
        input_source_mode = st.radio(
            "Nguồn đầu vào:",
            [
                "🎙️ File Ghi Âm (.webm / .wav / .mp3 / .m4a) - Hỗ trợ file lớn >75MB",
                "📄 File Phụ Đề (.txt) tải về từ Bước 1",
                "📝 Văn bản đã dịch / dán trực tiếp từ Bước 1"
            ]
        )

        uploaded_audio_recap = None
        uploaded_txt_recap = None
        text_for_recap = ""

        if input_source_mode.startswith("🎙️"):
            uploaded_audio_recap = st.file_uploader(
                "Tải lên File Ghi Âm cuộc họp (.webm, .wav, .mp3, .m4a):",
                type=["webm", "wav", "mp3", "m4a", "ogg"]
            )
            if uploaded_audio_recap:
                st.audio(uploaded_audio_recap)
        elif input_source_mode.startswith("📄"):
            uploaded_txt_recap = st.file_uploader("Tải lên File Phụ Đề (.txt):", type=["txt"])
            if uploaded_txt_recap:
                text_for_recap = uploaded_txt_recap.read().decode("utf-8", errors="ignore")
                st.text_area("Nội dung file .txt:", value=text_for_recap, height=110)
        else:
            combined_transcript = "\n".join(
                [f"[{x['time']}] {x['target']} (Gốc: {x['source']})" for x in st.session_state.transcript_list]
            ) or st.session_state.raw_stream_input
            text_for_recap = st.text_area("Nội dung văn bản cuộc họp:", value=combined_transcript, height=120)

    with col_prm:
        st.markdown("##### 📌 4. Kịch bản Prompt Recap (Tự động kết hợp Nhân sự & Mẫu BBH)")
        prompt_templates = {
            "1. Mẫu chuẩn Biên Bản Họp Phòng QLCL KCN - THACO AUTO (7 cột chuẩn)": """Đóng vai Trưởng Phòng Quản lý Chất lượng (Phòng QLCL KCN - THACO AUTO). Dựa vào toàn bộ nội dung cuộc họp, Danh sách Nhân sự THACO thực tế tham dự, Cấu trúc File Mẫu BBH và Bảng thuật ngữ THACO đính kèm, hãy phân tích và xuất ra JSON chuẩn để đổ trực tiếp vào mẫu Biên Bản Họp:
{
  "ve_viec": "Tóm tắt chủ đề cuộc họp ngắn gọn, chuẩn văn phong hành chính THACO (1 dòng để điền vào mục V/v: ...)",
  "rows": [
    {
      "tt": "1",
      "van_de": "Tên vấn đề / Hạng mục lỗi / Báo cáo khiếu nại cụ thể",
      "noi_dung_thao_luan": "Chi tiết nội dung các bên đã thảo luận, ý kiến của từng nhân sự/đơn vị tham gia, nguyên nhân gốc rễ, hiện trạng kỹ thuật",
      "cong_viec_tiep_theo": "Đối sách, công việc thực hiện tiếp theo cụ thể, rõ ràng",
      "trach_nhiem": "Đơn vị / Nhân sự chịu trách nhiệm chính (BẮT BUỘC đối chiếu đúng tên Phòng ban/Nhà máy/Nhân sự tham gia hoặc Đối tác)",
      "phoi_hop": "Đơn vị / Nhân sự phối hợp (Ghi '-' nếu không có)",
      "hoan_thanh": "Thời hạn hoàn thành (VD: 20/11/2026 hoặc Hằng tuần)"
    }
  ],
  "ket_luan_chung": "Câu kết luận chỉ đạo chung cuối biên bản"
}
Lưu ý: Chỉ trả về đúng chuỗi JSON hợp lệ, không thêm ký tự markdown ```json.""",

            "2. Họp Đàm phán Đối tác Quốc tế / Nhà cung cấp (BJM, Foton, Hyundai...)": """Đóng vai Chủ trì / Thư ký kỹ thuật Phòng QLCL KCN - THACO AUTO làm việc với đối tác quốc tế/NCC. Hãy đối chiếu Danh sách Nhân sự tham dự, Cấu trúc Mẫu BBH và Bảng thuật ngữ THACO để bóc tách từng vấn đề kỹ thuật, bảo hành, chất lượng xe thành JSON chuẩn 7 cột:
{
  "ve_viec": "Họp đánh giá chất lượng xe và giải quyết khiếu nại kỹ thuật với đối tác",
  "rows": [
    {
      "tt": "1",
      "van_de": "Mô tả vấn đề / Khiếu nại chất lượng / Mã lỗi",
      "noi_dung_thao_luan": "Ý kiến phản hồi của đối tác/NCC và đánh giá kỹ thuật của đại diện THACO tham dự họp",
      "cong_viec_tiep_theo": "Hành động khắc phục, cung cấp phụ tùng hoặc lập báo cáo kỹ thuật",
      "trach_nhiem": "Tên Đối tác / Đơn vị phụ trách chính",
      "phoi_hop": "Đơn vị phối hợp",
      "hoan_thanh": "Thời hạn hoàn thành"
    }
  ],
  "ket_luan_chung": "Trên đây là các thống nhất, chỉ đạo và triển khai trong cuộc họp, đề nghị các phòng ban, các Nhà máy triển khai thực hiện theo tiến độ, Phòng QLCL KCN làm đầu mối tổng hợp, theo dõi và báo cáo tiến độ thực hiện theo định kỳ."
}
Chỉ trả về đúng chuỗi JSON hợp lệ, không thêm ký tự markdown."""
        }

        selected_tpl = st.selectbox("Kịch bản Prompt:", list(prompt_templates.keys()), label_visibility="collapsed")
        custom_prompt = st.text_area("Nội dung Prompt chi tiết:", value=prompt_templates[selected_tpl], height=160)

    if st.button(
        f"✨ Chạy AI Recap Kết Hợp Danh Sách Nhân Sự Tham Dự & Form Mẫu BBH (`{selected_recap_model}`)",
        type="primary",
        use_container_width=True
    ):
        if not gemini_api_key:
            st.error("Vui lòng nhập Gemini API Key ở thanh bên trái!")
        elif input_source_mode.startswith("🎙️") and not uploaded_audio_recap:
            st.warning("Vui lòng tải lên File Ghi Âm!")
        elif not input_source_mode.startswith("🎙️") and not text_for_recap.strip():
            st.warning("Chưa có dữ liệu văn bản cuộc họp để phân tích!")
        else:
            with st.spinner("AI đang phân tích cuộc họp, đối chiếu Nhân sự tham gia, Cấu trúc Form BBH & Bảng thuật ngữ THACO..."):
                try:
                    template_ctx = (
                        f"--- CẤU TRÚC MẪU BIÊN BẢN HỌP (.DOCX) ĐẦU VÀO ---\n{st.session_state.template_structure_hint}\n\n"
                        if st.session_state.template_structure_hint else ""
                    )
                    shared_context = (
                        f"--- DANH SÁCH NHÂN SỰ THACO THỰC TẾ THAM GIA CUỘC HỌP NÀY ---\n"
                        f"{json.dumps(st.session_state.selected_attendees, ensure_ascii=False)}\n"
                        f"Nhân sự / Đối tác bổ sung: {st.session_state.extra_attendees_str}\n\n"
                        f"{template_ctx}"
                        f"--- BẢNG THUẬT NGỮ THACO DÙNG CHUNG ---\n"
                        f"{json.dumps(st.session_state.glossary, ensure_ascii=False)}\n\n"
                    )

                    if input_source_mode.startswith("🎙️"):
                        full_prompt = (
                            f"{custom_prompt}\n\n"
                            f"{shared_context}"
                            f"Hãy lắng nghe toàn bộ file ghi âm cuộc họp đính kèm, đối chiếu đúng nhân sự tham gia, thuật ngữ THACO và cấu trúc mẫu BBH để bóc tách chi tiết vào bảng JSON 7 cột."
                        )
                        raw_reply = generate_recap_with_audio_or_text(
                            gemini_api_key, selected_recap_model, full_prompt, audio_file=uploaded_audio_recap
                        )
                    else:
                        full_prompt = (
                            f"{custom_prompt}\n\n"
                            f"{shared_context}"
                            f"--- NỘI DUNG CUỘC HỌP ---\n{text_for_recap}"
                        )
                        raw_reply = generate_recap_with_audio_or_text(
                            gemini_api_key, selected_recap_model, full_prompt, audio_file=None
                        )

                    cleaned = raw_reply.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                    parsed = json.loads(cleaned)
                    
                    if parsed.get("ve_viec"):
                        st.session_state.bbh_meta["ve_viec"] = parsed["ve_viec"]
                    if parsed.get("ket_luan_chung"):
                        st.session_state.bbh_meta["ket_luan_chung"] = parsed["ket_luan_chung"]
                    if parsed.get("rows") and isinstance(parsed["rows"], list):
                        st.session_state.recap_rows = parsed["rows"]

                    st.success("✅ Đã phân tích thành công theo đúng Form mẫu & Nhân sự tham gia! Kiểm tra bảng bên dưới và bấm sang Bước 3 để tải Biên bản Word.")
                except Exception as e:
                    st.error(f"Lỗi phân tích AI: {e}")

    st.markdown("##### 🔍 Rà soát & Chỉnh sửa Bảng III: Nội dung thảo luận và kế hoạch triển khai (Liên thông sang Bước 3)")
    edited_rows = st.data_editor(
        st.session_state.recap_rows,
        column_config={
            "tt": st.column_config.TextColumn("TT", width="small"),
            "van_de": st.column_config.TextColumn("Vấn đề", width="medium"),
            "noi_dung_thao_luan": st.column_config.TextColumn("Nội dung thảo luận", width="large"),
            "cong_viec_tiep_theo": st.column_config.TextColumn("Công việc thực hiện tiếp theo", width="large"),
            "trach_nhiem": st.column_config.TextColumn("T.nhiệm", width="small"),
            "phoi_hop": st.column_config.TextColumn("P.hợp", width="small"),
            "hoan_thanh": st.column_config.TextColumn("H.thành", width="small"),
        },
        num_rows="dynamic",
        use_container_width=True,
        height=250,
        key="recap_table_editor"
    )
    st.session_state.recap_rows = edited_rows

# ----------------- TAB 3: XUẤT BIÊN BẢN HỌP WORD TRÌNH KÝ -----------------
with tab3:
    st.subheader("📄 Xuất Biên Bản Họp Word (.docx) Trình Ký Ngay")

    if st.session_state.template_docx_bytes is not None:
        st.success(
            f"✅ Hệ thống đã nhận sẵn File Mẫu BBH **`{st.session_state.template_docx_name}`** và **{len(st.session_state.selected_attendees)} nhân sự tham gia** từ Bước 2. Anh chỉ cần kiểm tra thông tin hành chính bên dưới và bấm Kết xuất!"
        )
    else:
        st.info("💡 Anh có thể tải File BBH Mẫu ở Bước 2 hoặc tải trực tiếp ở ô bên dưới. Nếu để trống, hệ thống sẽ tạo theo Form chuẩn của Phòng QLCL KCN - THACO AUTO.")

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.session_state.bbh_meta["so_bbh"] = st.text_input("Số Biên bản họp:", value=st.session_state.bbh_meta["so_bbh"])
        st.session_state.bbh_meta["ve_viec"] = st.text_input("V/v (Chủ đề cuộc họp):", value=st.session_state.bbh_meta["ve_viec"])
    with col_m2:
        st.session_state.bbh_meta["thoi_gian"] = st.text_input("I. Thời gian:", value=st.session_state.bbh_meta["thoi_gian"])
        st.session_state.bbh_meta["dia_diem"] = st.text_input("I. Địa điểm:", value=st.session_state.bbh_meta["dia_diem"])

    update_att_flag = st.checkbox(
        f"Tự động điền danh sách {len(st.session_state.selected_attendees)} nhân sự đã chọn ở Bước 2 vào Bảng II (Thành phần tham dự) trong file Word",
        value=True
    )

    st.session_state.bbh_meta["ket_luan_chung"] = st.text_area(
        "Câu kết luận chỉ đạo cuối biên bản:",
        value=st.session_state.bbh_meta["ket_luan_chung"],
        height=70
    )

    uploaded_docx_step3 = st.file_uploader(
        "📂 Tải lên File BBH Mẫu khác (Tùy chọn - nếu muốn thay đổi mẫu đã nạp ở Bước 2):",
        type=["docx"],
        key="uploader_tpl_step3"
    )
    if uploaded_docx_step3 is not None:
        st.session_state.template_docx_bytes = uploaded_docx_step3.getvalue()
        st.session_state.template_docx_name = uploaded_docx_step3.name

    if st.button("🖨️ Kết Xuất Biên Bản Họp Word (.docx) Để Trình Ký Ngay", type="primary", use_container_width=True):
        try:
            active_attendees = st.session_state.selected_attendees or st.session_state.personnel
            if st.session_state.template_docx_bytes is not None:
                final_docx_bytes = populate_uploaded_thaco_docx(
                    st.session_state.template_docx_bytes,
                    st.session_state.bbh_meta,
                    st.session_state.recap_rows,
                    active_attendees,
                    update_attendance_table=update_att_flag
                )
            else:
                final_docx_bytes = generate_default_thaco_bbh_docx(
                    st.session_state.bbh_meta,
                    st.session_state.recap_rows,
                    active_attendees
                )

            out_name = f"{datetime.now().strftime('%Y-%m%d')}. BB hop chat luong - QLCL KCN.docx"
            st.download_button(
                label=f"⬇️ TẢI XUỐNG BIÊN BẢN HỌP HOÀN CHỈNH ({out_name})",
                data=final_docx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
            st.success("✅ Đã xuất Biên bản họp đầy đủ Thành phần tham dự & Bảng nội dung 7 cột chuẩn Phòng QLCL KCN!")
        except Exception as e:
            st.error(f"Lỗi khi xuất file Word: {e}")

# ----------------- TAB 4: QUẢN TRỊ NHÂN SỰ, TM & GLOSSARY TRÊN GOOGLE SHEETS -----------------
with tab4:
    st.subheader("📚 Quản Trị Trực Tiếp Nhân Sự, Thuật Ngữ & Bộ Nhớ Dịch (Đồng Bộ Google Sheets)")

    st.markdown(f"### 👥 1. Danh Bạ Nhân Sự - Chức Danh - Phòng Ban (`Personnel`: {len(st.session_state.personnel)} nhân sự)")
    st.caption("Thêm/sửa/xóa dòng trực tiếp bên dưới rồi bấm Lưu. Danh sách này sẽ đồng bộ ngay với các nút 1-Click ở Bước 1 và danh sách chọn người tham dự ở Bước 2.")
    p_list = st.session_state.personnel if st.session_state.personnel else [{"ho_ten": "", "chuc_danh": "", "phong_ban": ""}]
    edited_personnel = st.data_editor(
        p_list,
        column_config={
            "ho_ten": st.column_config.TextColumn("Họ và Tên (VD: Ông Võ Thành Lưu)", width="medium"),
            "chuc_danh": st.column_config.TextColumn("Chức danh (VD: TP QLCL xe Thương mại)", width="medium"),
            "phong_ban": st.column_config.TextColumn("Phòng ban / Nhà máy (VD: Phòng QLCL KCN)", width="medium")
        },
        num_rows="dynamic",
        use_container_width=True,
        height=240,
        key="editor_personnel"
    )
    if st.button("💾 Lưu Danh sách Nhân sự lên Google Sheets (`save_personnel`)", type="primary", use_container_width=True):
        ok, msg = save_personnel_to_sheets(webapp_url, edited_personnel)
        if ok:
            st.success(msg)
        else:
            st.error(msg)

    st.divider()
    col_g, col_tm = st.columns(2)
    
    with col_g:
        st.markdown(f"**📖 2. Bảng Thuật Ngữ Chung (`Glossary`: {len(st.session_state.glossary)} từ)**")
        g_list = [{"Tu_Goc": k, "Tu_Dich": v} for k, v in st.session_state.glossary.items()] or [{"Tu_Goc": "", "Tu_Dich": ""}]
        edited_g = st.data_editor(g_list, num_rows="dynamic", use_container_width=True, height=300, key="editor_glossary")
        
        if st.button("💾 Lưu toàn bộ Bảng Thuật Ngữ lên Sheets (`save_glossary`)", use_container_width=True):
            new_g_dict = {row["Tu_Goc"].strip(): row["Tu_Dich"].strip() for row in edited_g if row.get("Tu_Goc") and row.get("Tu_Dich")}
            ok, msg = save_full_glossary_to_sheets(webapp_url, new_g_dict)
            if ok:
                st.success(msg)
            else:
                st.error(msg)

    with col_tm:
        st.markdown(f"**🧠 3. Bộ Nhớ Đệm Dịch Thuật Chung (`TM`: {len(st.session_state.tm)} câu)**")
        tm_list = [{"Khoa_Goc": k, "Ban_Dich_Chuan": v} for k, v in st.session_state.tm.items()] or [{"Khoa_Goc": "", "Ban_Dich_Chuan": ""}]
        edited_tm = st.data_editor(tm_list, num_rows="dynamic", use_container_width=True, height=300, key="editor_tm")
        
        if st.button("💾 Cập nhật các câu đã sửa lên Sheets (`upsert_tm`)", use_container_width=True):
            new_tm_dict = {row["Khoa_Goc"].strip(): row["Ban_Dich_Chuan"].strip() for row in edited_tm if row.get("Khoa_Goc") and row.get("Ban_Dich_Chuan")}
            ok, msg = upsert_tm_to_sheets(webapp_url, new_tm_dict)
            if ok:
                st.success(msg)
            else:
                st.error(msg)