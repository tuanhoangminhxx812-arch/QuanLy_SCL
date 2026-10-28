import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os, re, datetime
from io import BytesIO
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
from data_helpers import load_tonghop, load_pm092, load_gia_tri_hop_dong, load_capnhat_tiendo, load_pm092_monthly
from form_module import load_db_data, doc_so_vn
from cloud_export import (get_project_section, get_cost_breakdown,
    export_ttr_duyet_qt_word,
    export_tmqt_word, export_phieu_tham_tra_word,
    export_bao_cao_tham_tra_word, export_qd_phe_duyet_word, _safe_int, _fmt_money_dot, _format_date_vn)
from github_helper import gh_list_files, gh_upload_file, gh_delete_file, gh_upload_root_file, has_token
from thqt_parser import parse_thqt_excel, save_thqt_cache, load_thqt_cache

st.set_page_config(page_title="Quản lý Quyết toán SCL", layout="wide", page_icon="⚡")

try:
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ── CSS ──
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,.stApp{font-family:'Inter',sans-serif}
.material-symbols-rounded, [data-testid="stIconMaterial"], .stIcon { font-family: 'Material Symbols Rounded', sans-serif !important; }
.block-container{padding-top:1rem}
/* Sidebar Background: Deep Industrial Navy Gradient */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #091e3a 0%, #0f325e 35%, #14437d 70%, #1a569e 100%) !important;
    border-right: 1px solid rgba(255, 255, 255, 0.1);
}
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3,
[data-testid="stSidebar"] .stMarkdown h5,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stMarkdown span,
[data-testid="stSidebar"] label {color:#ffffff!important}

/* 3D Hộp Nối Chassis / Panel Frame */
[data-testid="stSidebar"] div[role="radiogroup"] {
    display: flex !important;
    flex-direction: column !important;
    gap: 12px !important;
    padding: 14px 10px !important;
    background: linear-gradient(180deg, rgba(5, 16, 32, 0.7) 0%, rgba(10, 26, 50, 0.55) 100%) !important;
    border-radius: 16px !important;
    border-top: 1.5px solid rgba(255, 255, 255, 0.15) !important;
    border-left: 1px solid rgba(255, 255, 255, 0.08) !important;
    border-right: 1px solid rgba(0, 0, 0, 0.3) !important;
    border-bottom: 3px solid rgba(0, 0, 0, 0.6) !important;
    box-shadow: inset 0 3px 12px rgba(0, 0, 0, 0.6), 0 2px 8px rgba(0, 0, 0, 0.25) !important;
}

/* 3D Modular Block (Hộp Nối 3D Thiết bị) */
[data-testid="stSidebar"] div[role="radiogroup"] > label {
    position: relative !important;
    background: linear-gradient(180deg, #1b4982 0%, #133966 50%, #0c2747 100%) !important;
    border-top: 1.5px solid rgba(255, 255, 255, 0.35) !important;
    border-left: 1.5px solid rgba(255, 255, 255, 0.15) !important;
    border-right: 1.5px solid rgba(0, 0, 0, 0.35) !important;
    border-bottom: 4px solid #07172b !important;
    border-radius: 11px !important;
    padding: 12px 14px !important;
    cursor: pointer !important;
    transition: all 0.16s cubic-bezier(0.4, 0, 0.2, 1) !important;
    display: flex !important;
    align-items: center !important;
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.35), inset 0 1px 0 rgba(255, 255, 255, 0.22), inset 0 -1px 0 rgba(0, 0, 0, 0.4) !important;
}

/* Typography inside 3D Box */
[data-testid="stSidebar"] div[role="radiogroup"] > label p,
[data-testid="stSidebar"] div[role="radiogroup"] > label span {
    color: #e2e8f0 !important;
    font-weight: 700 !important;
    font-size: 12px !important;
    letter-spacing: 0.3px !important;
    line-height: 1.35 !important;
    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.6) !important;
    margin: 0 !important;
}

/* 3D Hover Effect */
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
    background: linear-gradient(180deg, #2563eb 0%, #1d4ed8 50%, #163e8c 100%) !important;
    transform: translateY(-2px) !important;
    border-top: 1.5px solid rgba(255, 255, 255, 0.65) !important;
    border-bottom: 5px solid #091a35 !important;
    box-shadow: 0 8px 18px rgba(0, 0, 0, 0.45), 0 0 15px rgba(37, 99, 235, 0.4), inset 0 1px 0 rgba(255, 255, 255, 0.45) !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label:hover p {
    color: #ffffff !important;
    text-shadow: 0 1px 3px rgba(0, 0, 0, 0.8) !important;
}

/* 3D Press Effect */
[data-testid="stSidebar"] div[role="radiogroup"] > label:active {
    transform: translateY(2px) !important;
    border-bottom: 1.5px solid #07172b !important;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.4), inset 0 2px 5px rgba(0, 0, 0, 0.5) !important;
}

/* 3D Active / Selected State (Khối hộp nổi bật / Đèn nguồn Amber sáng) */
[data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"] {
    background: linear-gradient(180deg, #ffffff 0%, #f8fafc 50%, #e2e8f0 100%) !important;
    border-top: 2px solid #ffffff !important;
    border-left: 6px solid #f59e0b !important;
    border-right: 1.5px solid #cbd5e1 !important;
    border-bottom: 4px solid #94a3b8 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 16px rgba(0, 0, 0, 0.35), 0 0 16px rgba(245, 158, 11, 0.35), inset 0 2px 0 #ffffff, inset 0 -1px 0 rgba(0, 0, 0, 0.1) !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"] p,
[data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"] span {
    color: #0b2246 !important;
    font-weight: 800 !important;
    text-shadow: none !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] div[data-testid="stMarkdownContainer"] {
    margin-left: 4px;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {
    display: none;
}
.main-title {
    text-align: center;
    color: #1565C0;
    margin-top: -40px;
    margin-bottom: 25px;
    font-size: 32px;
    font-weight: 700;
    letter-spacing: 1px;
}
/* Sticky header: pin the header container when scrolling */
div[data-testid="stVerticalBlock"]:has(.header-content-wrapper) {
    position: sticky !important;
    top: 0 !important;
    z-index: 999 !important;
    background-color: white !important;
    padding: 10px 0 10px 0;
    margin-top: -15px;
    border-bottom: 2px solid #e2e8f0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}
.metric-card{background:linear-gradient(135deg,#1565C0,#1E88E5);border-radius:14px;
padding:20px 16px;text-align:center;border:1px solid rgba(255,255,255,.2);
box-shadow:0 4px 20px rgba(21,101,192,.3);transition:transform .2s}
.metric-card:hover{transform:translateY(-3px)}
.metric-val{font-size:26px;font-weight:700;color:#ffffff}
.metric-label{font-size:12px;color:rgba(255,255,255,.8);margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px}
.page-title{font-size:28px;font-weight:700;color:#1565C0;margin-bottom:8px}
.section-title{font-size:18px;font-weight:600;color:#1565C0;margin:16px 0 8px;
border-left:3px solid #1976D2;padding-left:10px}
.sidebar-logo{font-size:32px!important;font-weight:800!important;color:#ffffff!important;
letter-spacing:1px;text-shadow:0 2px 8px rgba(0,0,0,.2)}
.sidebar-sub{font-size:16px!important;color:rgba(255,255,255,.85)!important;font-weight:500!important}
.a4-preview{background:#ffffff;border:1px solid #d0d0d0;padding:2cm 2.5cm;
max-height:700px;overflow-y:auto;font-family:'Times New Roman',serif;font-size:13pt;
line-height:1.6;color:#000000;box-shadow:0 2px 20px rgba(0,0,0,.18);margin:0 auto;max-width:210mm;border-radius:2px;}
.a4-preview p{margin:0 0 4pt 0;text-align:justify;}
.a4-preview .center-title{text-align:center;font-weight:bold;font-size:14pt;margin-bottom:12pt;line-height:1.3;text-transform:uppercase;}
.a4-preview table{width:100%;border-collapse:collapse;margin:8pt 0;}
.a4-preview td,.a4-preview th{padding:5pt;vertical-align:top;}
div[data-testid="stDataFrame"]{border-radius:10px;overflow:hidden}
.doc-card{background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:16px 18px;margin-bottom:4px;}
.doc-card-title{font-weight:700;color:#1565C0;font-size:14px;margin-bottom:4px;display:flex;align-items:center;gap:8px;}
.doc-card-sub{font-size:12px;color:#64748b;margin-bottom:8px;}
.doc-file-chip{display:inline-flex;align-items:center;gap:5px;background:#e0edff;color:#1565C0;border-radius:20px;padding:3px 10px;font-size:12px;margin:2px;}
/* Custom HTML table for Bảng tổng hợp QT */
.qt-table{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:12px;margin-top:8px;}
.qt-table thead th{
    background:linear-gradient(135deg,#1565C0,#1976D2);color:#fff;font-weight:600;
    padding:8px 8px;text-align:center;border:1px solid #1255a0;
    position:sticky;top:0;z-index:10;font-size:11px;letter-spacing:0.3px;
}
.qt-table tbody td{
    padding:6px 8px;border:1px solid #e2e8f0;vertical-align:middle;
}
.qt-table tbody tr:nth-child(even){background:#f8fafc;}
.qt-table tbody tr:hover{background:#e8f0fe;}
.qt-table .col-stt{width:40px;text-align:center;font-weight:600;color:#1565C0;white-space:nowrap;}
.qt-table .col-ten{text-align:left;min-width:180px;font-size:12px;}
.qt-table .col-num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;min-width:90px;}
.qt-table .row-header td{font-weight:700;background:#eef2ff !important;}
.qt-table .row-total td{font-weight:700;background:#dbeafe !important;border-top:2px solid #1565C0;}
.qt-table-wrapper{max-height:700px;overflow-y:auto;border-radius:10px;border:1px solid #e2e8f0;box-shadow:0 2px 8px rgba(0,0,0,0.06);}
</style>""", unsafe_allow_html=True)

# ── Helpers ──
def create_download_link(data, filename, btn_label, mime):
    import base64, uuid
    b64 = base64.b64encode(data).decode()
    uid = "dl_" + str(uuid.uuid4()).replace('-', '')
    css = f"""
    <style>
    #{uid} {{
        display: inline-flex; align-items: center; justify-content: center;
        background-color: #ffffff; color: #31333F; border: 1px solid rgba(49,51,63,0.2);
        border-radius: 0.5rem; padding: 0.5rem 1rem; font-size: 1rem; font-weight: 400;
        text-decoration: none; transition: all 0.2s; cursor: pointer; margin-top: 10px;
    }}
    #{uid}:hover {{ border-color: #ff4b4b; color: #ff4b4b; }}
    </style>
    """
    return f'{css}<a id="{uid}" href="data:{mime};base64,{b64}" download="{filename}">{btn_label}</a>'

def clean_filename(name):
    if not name: return "CT"
    s1 = u'ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúýĂăĐđĨĩŨũƠơƯưẠạẢảẤấẦầẨẩẪẫẬậẮắẰằẲẳẴẵẶặẸẹẺẻẼẽẾếỀềỂểỄễỆệỈỉỊịỌọỎỏỐốỒồỔổỖỗỘộỚớỜờỞởỠỡỢợỤụỦủỨứỪừỬửỮữỰựỲỳỴỵỶỷỸỹ'
    s0 = u'AAAAEEEIIOOOOUUYaaaaeeeiioooouuyAaDdIiUuOoUuAaAaAaAaAaAaAaAaAaAaAaAaEeEeEeEeEeEeEeEeIiIiOoOoOoOoOoOoOoOoOoOoOoOoUuUuUuUuUuUuUuYyYyYyYy'
    s = ''
    for c in str(name):
        if c in s1: s += s0[s1.index(c)]
        else: s += c
    import re
    return re.sub(r'[^a-zA-Z0-9_\-\. ]', '_', s)[:30].strip(' _')

def fmt_money(v):
    try:
        v=float(v)
        if v>=1e9: s = f"{v/1e9:,.2f} tỷ"
        elif v>=1e6: s = f"{v/1e6:,.1f} tr"
        else:
            if v.is_integer(): s = f"{int(v):,}"
            else: s = f"{v:,.1f}"
        return s.replace(',', 'X').replace('.', ',').replace('X', '.')
    except:return "0"

def fmt_full(v):
    try:
        f = float(v)
        if f.is_integer(): s = f"{int(f):,}"
        else: s = f"{f:,.1f}"
        return s.replace(',', 'X').replace('.', ',').replace('X', '.')
    except:return "0"

# ── Helper: wrap HTML preview với full CSS để components.html render đúng ──
A4_CSS = """
<style>
*{box-sizing:border-box}
body{margin:0;padding:12px;background:#f0f2f6;font-family:'Times New Roman',serif;font-size:13pt;color:#000}
.a4-preview{background:#ffffff;border:1px solid #d0d0d0;padding:2cm 2.5cm;
 max-width:210mm;margin:0 auto;line-height:1.6;box-shadow:0 2px 20px rgba(0,0,0,.18);border-radius:2px;}
.a4-preview p{margin:0 0 4pt 0;text-align:justify;}
.a4-preview table{width:100%;border-collapse:collapse;margin:8pt 0;}
.a4-preview td,.a4-preview th{padding:5pt;vertical-align:top;}
/* Bảng số liệu chi phí ngay hàng thẳng lối 100% chuẩn mẫu */
.cost-table { width: auto !important; border: none !important; margin: 4pt 0 6pt 35px !important; border-collapse: collapse !important; }
.cost-table tr { border: none !important; }
.cost-table td { border: none !important; padding: 1.5pt 0 !important; vertical-align: middle !important; font-family: 'Times New Roman', serif !important; font-size: 13pt !important; }
.cost-table td.col-lbl { white-space: nowrap !important; text-align: left !important; }
.cost-table td.col-colon { width: 24px !important; text-align: center !important; padding: 1.5pt 12px !important; }
.cost-table td.col-val { text-align: right !important; white-space: nowrap !important; font-variant-numeric: tabular-nums !important; }
.cost-table td.col-unit { text-align: left !important; white-space: nowrap !important; padding-left: 8px !important; width: 60px !important; }
</style>
"""
def wrap_preview_html(body_html):
    return f"<!DOCTYPE html><html><head><meta charset='utf-8'>{A4_CSS}</head><body>{body_html}</body></html>"

def _render_cost_table_html(items, indent_px=35, label_min_width=210):
    """
    Hiển thị khối số liệu chi phí 'ngay hàng thẳng lối' 100% chuẩn đẹp chuyên nghiệp:
    - Nhãn căn lề trái (min-width cố định để colon gióng thẳng tắp)
    - Dấu hai chấm (:) căn giữa hoàn toàn trong ô riêng
    - Số tiền căn lề phải (tabular-nums để các chữ số gióng thẳng tắp)
    - Đơn vị tính (đồng. / đồng) thẳng hàng ngay bên phải số tiền
    """
    rows = []
    for item in items:
        lbl, val, is_bold = item[0], item[1], item[2]
        unit = item[3] if len(item) > 3 else "đồng."
        b_s = "<b>" if is_bold else ""
        b_e = "</b>" if is_bold else ""
        rows.append(
            f'<tr>'
            f'<td class="col-lbl" style="min-width:{label_min_width}px;">{b_s}{lbl}{b_e}</td>'
            f'<td class="col-colon">{b_s}:{b_e}</td>'
            f'<td class="col-val">{b_s}{val}{b_e}</td>'
            f'<td class="col-unit">{b_s}{unit.strip()}{b_e}</td>'
            f'</tr>'
        )
    return (
        f'<table class="cost-table" style="margin-left:{indent_px}px !important;">'
        + "".join(rows)
        + '</table>'
    )

def status_color(s):
    m={'Đang thi công':'#22c55e','Lập PAKT-Tổng dự toán':'#f59e0b',
       'Lập kế hoạch đấu thầu':'#3b82f6','Hoàn thành':'#8b5cf6','Nghiệm thu':'#06b6d4'}
    return m.get(str(s).strip(),'#94a3b8')


def parse_date_from_text(text, prefix):
    m=re.search(fr"{prefix}:\s*(\d{{1,2}})/(\d{{4}})",str(text),re.IGNORECASE)
    return (int(m.group(1)),int(m.group(2))) if m else (None,None)

def analyze_health(row,cy,cm):
    status=str(row.get('Trạng thái','')).strip()
    td=str(row.get('Tiến độ',''))
    kt=float(row.get('Khái toán',0)) if pd.notna(row.get('Khái toán')) else 0
    th=float(row.get('Thực hiện',0)) if pd.notna(row.get('Thực hiện')) else 0
    r=(th/kt*100) if kt>0 else 0
    
    if status in['Hoàn thành','Nghiệm thu']:return "Hoàn thành","🟢","Đã hoàn thành", r
    kc_m,kc_y=parse_date_from_text(td,'KC');ht_m,ht_y=parse_date_from_text(td,'HT')
    if ht_m and ht_y:
        ml=(ht_y-cy)*12+(ht_m-cm)
        if ml<0:return "Quá hạn","🔴",f"Trễ {-ml} tháng (HT: {ht_m}/{ht_y})", r
        if ml<=2 and r<30:return "Nguy cơ","🟡",f"Còn {ml} tháng, GN thấp", r
    if kc_m and kc_y:
        mp=(cy-kc_y)*12+(cm-kc_m)
        if mp>2 and status in['Lập PAKT-Tổng dự toán','Lập kế hoạch đấu thầu']:
            return "Trễ tiến độ","🔴",f"Qua KC {mp} tháng, vẫn '{status}'", r
    return "Bình thường","🔵","Tiến độ BT", r

# ── Load data ──
@st.cache_data(ttl=300)
def load_all():
    df=load_tonghop();pm=load_pm092();hd=load_gia_tri_hop_dong()
    if not df.empty and 'Mã CT' in df.columns:
        df['Thực hiện PM']=df['Mã CT'].map(lambda x:pm.get(str(x).strip(),0))
        if 'Thực hiện' not in df.columns:df['Thực hiện']=df['Thực hiện PM']
        else:df['Thực hiện']=df.apply(lambda r:r['Thực hiện PM'] if r['Thực hiện PM']>0 else r['Thực hiện'],axis=1)
        df['Giá trị HĐ']=df['Mã CT'].map(lambda x:hd.get(str(x).strip(),0))
    return df

df_th=load_all()
db_df=load_db_data()

# ── Session state ──
if 'dli_files' not in st.session_state:
    st.session_state.dli_files = {}  # {ma_ct: {doc_type: [file_name_list]}}

# ── Sidebar ──
with st.sidebar:
    st.markdown('<p class="sidebar-logo" style="font-size: 28px !important; line-height: 1.3; margin-bottom: 5px;">Công ty Điện lực Vũng Tàu</p>',unsafe_allow_html=True)
    st.divider()
    page=st.radio("📂 Chuyên mục",
        [
            "📊 TỔNG QUAN",
            "📁 KIỂM TRA HỒ SƠ",
            "📋 BẢNG TỔNG HỢP QUYẾT TOÁN",
            "📝 TỜ TRÌNH DUYỆT QUYẾT TOÁN DANH MỤC SCL",
            "📜 BÁO CÁO THẨM TRA",
            "⚖️ QUYẾT ĐỊNH PHÊ DUYỆT QUYẾT TOÁN"
        ],
        label_visibility="collapsed")
    st.divider()
    st.caption(f"Cập nhật: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}")



# ── Chọn CT helper ──
def _select_ct(key, container=None):
    if container is None: container = st
    names = []
    default_idx = 0
    cur_selected_ma = st.session_state.get('selected_ma_ct', '')
    for idx, (_, r) in enumerate(df_th.iterrows()):
        ma = str(r.get('Mã CT', '')).strip()
        ten = str(r.get('Tên công trình', '')).strip()
        names.append(f"{ma} - {ten}")
        if cur_selected_ma and ma == cur_selected_ma:
            default_idx = idx

    sel = container.selectbox("Chọn công trình:", names, index=default_idx, key=key)
    if not sel: return None, None, None, pd.DataFrame()
    ma = sel.split(" - ")[0].strip()
    st.session_state['selected_ma_ct'] = ma
    ri = df_th[df_th['Mã CT'] == ma].index
    if len(ri) == 0: return None, None, None, pd.DataFrame()
    row_th = df_th.loc[ri[0]]
    ten = str(row_th.get('Tên công trình', ''))
    mr, cd = get_project_section(db_df, ten, ma) if not db_df.empty else (None, pd.DataFrame())
    return row_th, mr, ten, cd

def shorten_name(name):
    n = str(name).lower()
    if 'đường dây trung' in n and 'hạ thế' in n: return 'Công trình SCL Đường dây trung-Hạ thế'
    if 'tu' in n and 'ti' in n and 'bảo trì' in n: return 'Công trình SCL Bảo trì TU_TI'
    if 'fco' in n or 'lbfco' in n: return 'Công trình SCL Thay thế FCO, LBFCO, LA'
    if 'cummins' in n: return 'Công trình SCL Máy phát điện Cummins G1, G2'
    if 'máy phát điện' in n and 'an hội' in n: return 'Công trình SCL Máy phát điện_Nhà máy An Hội'
    if 'công xa' in n: return 'Công trình SCL Công Xa'
    if 'live-line' in n or 'live line' in n: return 'Công trình SCL HTĐ bằng Live-line'
    return name

# ── GLOBAL HEADER ──
st.markdown('<div class="header-content-wrapper"></div>', unsafe_allow_html=True)
st.markdown('<h1 class="main-title">HỆ THỐNG QUẢN LÝ QUYẾT TOÁN SỬA CHỮA LỚN</h1>', unsafe_allow_html=True)

# ── PAGE 1: TỔNG QUAN ──
if "TỔNG QUAN" in page:
    if df_th.empty:
        st.warning("Chưa có dữ liệu Tổng hợp.xlsx")
    else:
        df_td = load_capnhat_tiendo()
        pm_monthly_all = load_pm092_monthly()
        
        if not df_td.empty:
            total_ct = len(df_td)
            total_kh = int(df_td['Khái toán'].fillna(0).sum()) if 'Khái toán' in df_td.columns else 0
        else:
            total_ct = len(df_th)
            total_kh = int(df_th['Khái toán'].fillna(0).sum()) if 'Khái toán' in df_th.columns else 0
            
        total_th = sum(sum(m_data.values()) for m_data in pm_monthly_all.values())
        ty_le = (total_th/total_kh*100) if total_kh>0 else 0

        c1, c2, c3, c4 = st.columns(4)
        for col,lbl,val in [(c1,"TỔNG SỐ CÔNG TRÌNH",str(total_ct)),
            (c2,"TỔNG KHÁI TOÁN",fmt_money(total_kh)),
            (c3,"TỔNG THỰC HIỆN",fmt_money(total_th)),
            (c4,"TỶ LỆ GIẢI NGÂN",f"{ty_le:.1f}%")]:
            col.markdown(f'<div class="metric-card"><div class="metric-label">{lbl}</div><div class="metric-val">{val}</div></div>',unsafe_allow_html=True)

        # ── BẢNG A: CẬP NHẬT TIẾN ĐỘ ──
        st.markdown('<p class="section-title">📋 Cập nhật tiến độ thực hiện SCL</p>',unsafe_allow_html=True)
        # df_td was already loaded above
        if not df_td.empty:
            status_badges = {
                'Đang thi công': '#22c55e',
                'Lập PAKT-Tổng dự toán': '#f59e0b',
                'Lập kế hoạch đấu thầu': '#3b82f6',
                'Hoàn thành': '#6366f1',
            }
            td_rows = ''
            for _, r in df_td.iterrows():
                tt = r.get('Trạng thái', '')
                badge_color = status_badges.get(tt, '#94a3b8')
                badge = f'<span style="background:{badge_color};color:#fff;padding:3px 8px;border-radius:12px;font-size:11px;white-space:nowrap;">{tt}</span>'
                kh_text = str(r.get('Tiến độ KH','')).replace('\\n','<br>').replace('\n','<br>')
                
                # Format Đã thực hiện as % if numeric
                thuc_hien_val = r.get('Đã thực hiện', '')
                
                try:
                    # Attempt to parse as float (e.g. 0.25)
                    float_val = float(thuc_hien_val)
                    if pd.notna(float_val):
                        thuc_hien_val = f"{float_val:.0%}"
                    else:
                        thuc_hien_val = ""
                except (ValueError, TypeError):
                    thuc_hien_val = str(thuc_hien_val)
                    
                td_rows += f'''<tr>
                    <td class="col-stt">{r.get('STT','')}</td>
                    <td style="text-align:center;font-weight:600;color:#1565C0;white-space:nowrap;font-size:11px;">{r.get('Mã CT','')}</td>
                    <td class="col-ten">{r.get('Tên công trình','')}</td>
                    <td style="font-size:11px;white-space:nowrap;line-height:1.4;">{kh_text}</td>
                    <td class="col-num">{fmt_full(r.get('Khái toán',0))}</td>
                    <td style="text-align:center;">{badge}</td>
                    <td style="font-size:12px;">{thuc_hien_val}</td>
                    <td style="font-size:11px;color:#64748b;">{r.get('Ghi chú','')}</td>
                </tr>\n'''
            td_html = f'''<div class="qt-table-wrapper" style="max-height:500px;overflow-x:auto;">
            <table class="qt-table" style="min-width:900px;">
                <thead><tr>
                    <th style="width:30px;">STT</th>
                    <th style="min-width:90px;">Mã CT</th>
                    <th style="min-width:180px;">Tên công trình</th>
                    <th style="min-width:110px;">Tiến độ KH</th>
                    <th style="min-width:100px;">Khái toán</th>
                    <th style="min-width:100px;">Trạng thái</th>
                    <th style="min-width:120px;">Đã thực hiện</th>
                    <th style="min-width:100px;">Ghi chú</th>
                </tr></thead>
                <tbody>{td_rows}</tbody>
            </table></div>'''
            st.markdown(td_html, unsafe_allow_html=True)
        else:
            st.info("Chưa có file CapNhatTienDo.xlsx")

        # ── BẢNG B: GIẢI NGÂN THEO THÁNG ──
        st.markdown('<p class="section-title">💰 Giải ngân theo tháng (6 tháng đầu năm)</p>',unsafe_allow_html=True)
        pm_monthly = pm_monthly_all # Already loaded above
        months = [1,2,3,4,5,6]
        month_labels = ['T1','T2','T3','T4','T5','T6']
        
        # Build header
        gn_header = '<th style="width:30px;">STT</th><th style="min-width:90px;">Mã CT</th><th style="min-width:160px;">Tên công trình</th><th style="min-width:80px;">GT HĐ</th>'
        for ml in month_labels:
            gn_header += f'<th style="min-width:90px;">{ml}</th>'
        gn_header += '<th style="min-width:100px;">Tổng GN</th>'
        
        # Build rows from df_th
        gn_rows = ''
        total_by_month = {m: 0 for m in months}
        grand_total = 0
        for idx, (_, r) in enumerate(df_th.iterrows()):
            ma = str(r.get('Mã CT','')).strip()
            ten = shorten_name(r.get('Tên công trình',''))
            hd_val = r.get('Giá trị HĐ', 0)
            try:
                hd_val = int(float(hd_val)) if pd.notna(hd_val) else 0
            except:
                hd_val = 0
            monthly_data = pm_monthly.get(ma, {})
            row_total = sum(monthly_data.get(m, 0) for m in months)
            grand_total += row_total
            
            gn_rows += f'<tr><td class="col-stt">{idx+1}</td>'
            gn_rows += f'<td style="text-align:center;font-weight:600;color:#1565C0;white-space:nowrap;font-size:11px;">{ma}</td>'
            gn_rows += f'<td class="col-ten">{ten}</td>'
            gn_rows += f'<td class="col-num">{fmt_full(hd_val) if hd_val else "0"}</td>'
            for m in months:
                val = monthly_data.get(m, 0)
                total_by_month[m] += val
                cell_style = 'class="col-num"'
                if val > 0:
                    cell_style = 'class="col-num" style="color:#1565C0;font-weight:600;"'
                gn_rows += f'<td {cell_style}>{fmt_full(val) if val else "-"}</td>'
            gn_rows += f'<td class="col-num" style="font-weight:700;color:#1565C0;">{fmt_full(row_total) if row_total else "0"}</td>'
            gn_rows += '</tr>\n'
        
        # Total row
        gn_rows += '<tr class="row-total"><td></td><td></td><td style="font-weight:700;font-size:12px;">TỔNG CỘNG</td>'
        gn_rows += f'<td></td>'
        for m in months:
            gn_rows += f'<td class="col-num">{fmt_full(total_by_month[m]) if total_by_month[m] else "-"}</td>'
        gn_rows += f'<td class="col-num" style="font-weight:700;color:#1565C0;">{fmt_full(grand_total)}</td></tr>'
        
        gn_html = f'''<div class="qt-table-wrapper" style="max-height:500px;overflow-x:auto;">
        <table class="qt-table" style="min-width:1000px;">
            <thead><tr>{gn_header}</tr></thead>
            <tbody>{gn_rows}</tbody>
        </table></div>'''
        st.markdown(gn_html, unsafe_allow_html=True)

        # ── BIỂU ĐỒ (đã chuyển xuống cuối) ──
        if HAS_PLOTLY:
            st.markdown('<p class="section-title">📈 Biểu đồ trực quan</p>',unsafe_allow_html=True)
            ch1,ch2=st.columns(2)
            with ch1:
                # Use df_td for charts if available
                df_chart = df_td if not df_td.empty else df_th
                if 'Trạng thái' in df_chart.columns:
                    sc=df_chart['Trạng thái'].fillna('Chưa XĐ').value_counts()
                    cm2={s:status_color(s) for s in sc.index}
                    fig1=px.pie(values=sc.values,names=sc.index,title="Tỷ trọng trạng thái",
                        color=sc.index,color_discrete_map=cm2,hole=0.4)
                    fig1.update_layout(paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='white',size=12),height=380,margin=dict(t=40,b=20,l=20,r=20))
                    fig1.update_traces(textposition='outside',textinfo='percent+label',textfont_size=11)
                    st.plotly_chart(fig1,width='stretch',key='pie1')
            with ch2:
                if 'Khái toán' in df_chart.columns:
                    top=df_chart.nlargest(5,'Khái toán').copy()
                    top['KT']=top['Khái toán'].fillna(0)/1e9
                    
                    # Calculate TH from pm_monthly_all for these top projects
                    th_values = []
                    for ma in top['Mã CT'] if 'Mã CT' in top.columns else []:
                        val = sum(pm_monthly_all.get(str(ma).strip(), {}).values())
                        th_values.append(val/1e9)
                    top['TH'] = th_values
                    
                    fig2=go.Figure()
                    fig2.add_trace(go.Bar(name='Khái toán (Tỷ)',x=top['Mã CT'] if 'Mã CT' in top.columns else top.index,y=top['KT'],marker_color='#6366f1'))
                    fig2.add_trace(go.Bar(name='Thực hiện (Tỷ)',x=top['Mã CT'] if 'Mã CT' in top.columns else top.index,y=top['TH'],marker_color='#f59e0b'))
                    fig2.update_layout(title="Top 5 ngân sách",barmode='group',
                        paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='white',size=12),height=380,margin=dict(t=40,b=20,l=20,r=20))
                    st.plotly_chart(fig2,width='stretch',key='bar1')





# ── PAGE 2: KIỂM TRA HỒ SƠ & PHÁP LÝ ──
elif "KIỂM TRA HỒ SƠ" in page:
    st.markdown('<p class="page-title">📁 KIỂM TRA HỒ SƠ & BỘ LỌC PHÁP LÝ THANH QUYẾT TOÁN SCL</p>', unsafe_allow_html=True)
    st.caption("Căn cứ: QĐ 202/QĐ-HĐTV (EVNHCMC) • TTr 1093/KTAT • Thông báo 2902/TB-PCVT • Nghị định 123/2020/NĐ-CP • Luật Đấu thầu 2023")
    
    row_th, mr, ten, cd = _select_ct("global_project_sel")
    if row_th is not None:
        ma_ct = str(row_th.get('Mã CT', '')).strip()
        from dossier_scanner import find_project_folder, scan_and_update_project, load_cached_scan_result
        from legal_checker import audit_dossier_legal
        
        p_folder = find_project_folder(ma_ct, ten)
        
        # Thanh công cụ quét & kiểm tra
        with st.container(border=True):
            col_f1, col_f2 = st.columns([3, 1])
            with col_f1:
                if p_folder and os.path.exists(p_folder):
                    rel_p = os.path.relpath(p_folder, BASE_DIR)
                    st.markdown(f"📁 **Thư mục hồ sơ:** `{rel_p}` ✅")
                    st.write(f"Công trình: **{ma_ct} - {ten}**")
                else:
                    st.markdown(f"📁 **Thư mục hồ sơ:** Chưa tìm thấy thư mục cho **{ma_ct}** ⚠️")
                    st.caption("💡 Chép tài liệu vào thư mục `Ho_so_cong_trinh/` để hệ thống tự động kiểm dò.")
            with col_f2:
                btn_audit = st.button("🚀 Quét & Kiểm dò Pháp lý", type="primary", use_container_width=True, key=f"btn_audit_{ma_ct}")

        if btn_audit:
            with st.spinner("🔄 Đang chạy Bộ lọc kiểm dò 7 tiêu chí pháp lý cốt lõi..."):
                res_scan = scan_and_update_project(ma_ct, ten)
                st.session_state[f"scan_res_{ma_ct}"] = res_scan
                st.rerun()

        # Nạp dữ liệu kiểm tra
        res_scan = st.session_state.get(f"scan_res_{ma_ct}")
        if not res_scan:
            res_scan = load_cached_scan_result(ma_ct)

        scanned_files = res_scan.get('scanned_files', []) if res_scan else []
        if not scanned_files and p_folder and os.path.exists(p_folder):
            for r_p, _, f_list in os.walk(p_folder):
                for f in f_list:
                    scanned_files.append(os.path.join(r_p, f))

        audit_res = audit_dossier_legal(ma_ct, ten, scanned_files, mr if mr is not None else row_th, cd, p_folder or "")

        # ── 1. BANNER KẾT LUẬN THANH TOÁN ──
        if not audit_res['can_pay'] or audit_res['overall_status'] == 'DANGER':
            st.error(f"""
            ### 🛑 KẾT LUẬN: CHƯA ĐỦ ĐIỀU KIỆN THANH TOÁN / KẾT THÚC HỢP ĐỒNG!
            **Lý do:** Hồ sơ còn tồn tại các lỗi pháp lý nghiêm trọng (màu đỏ bên dưới) cần khắc phục hoàn tất trước khi trình Giám đốc phê duyệt chi tiền.
            """)
        elif audit_res['overall_status'] == 'WARNING':
            st.warning(f"""
            ### ⚠️ KẾT LUẬN: CẦN BỔ SUNG & RÀ SOÁT CHỨNG TỪ
            Hồ sơ cơ bản đáp ứng, nhưng cần bổ sung các tài liệu còn thiếu để đảm bảo an toàn tuyệt đối khi thanh quyết toán.
            """)
        else:
            st.success(f"""
            ### 🟢 KẾT LUẬN: ĐỦ ĐIỀU KIỆN THANH QUYẾT TOÁN HỢP ĐỒNG!
            Hồ sơ tuân thủ 100% các quy định pháp lý, tên đơn vị, địa chỉ, tài khoản ngân hàng, chứng từ chất lượng CO/CQ và hóa đơn.
            """)

        # ── 2. KHUNG CÁC VIỆC CẦN XỬ LÝ NGAY ──
        if audit_res['action_items']:
            with st.container(border=True):
                st.markdown("### 📋 DANH SÁCH VIỆC CẦN XỬ LÝ NGAY ĐỂ HOÀN TẤT HỒ SƠ:")
                st.caption("Anh hãy kiểm tra và đánh dấu các công việc đã thực hiện xong bên dưới:")
                for idx_act, item in enumerate(audit_res['action_items']):
                    st.checkbox(f"**{item}**", key=f"act_{ma_ct}_{idx_act}")

        # ── 3. BẢNG CHI TIẾT 7 TIÊU CHÍ KIỂM DÒ PHÁP LÝ ──
        st.markdown("### 🔍 BẢNG CHI TIẾT 7 TIÊU CHÍ KIỂM DÒ PHÁP LÝ:")
        for chk in audit_res['checks']:
            icon = chk.get('icon', '⚪')
            lvl = chk.get('level', 'OK')
            border_color = "#ef4444" if lvl == "DANGER" else ("#f59e0b" if lvl == "WARNING" else "#22c55e")
            
            with st.container(border=True):
                c_title, c_stat = st.columns([3, 1])
                with c_title:
                    st.markdown(f"#### {icon} {chk['category']}: **{chk['title']}**")
                    st.caption(f"📌 **Căn cứ pháp lý:** `{chk['rule']}`")
                with c_stat:
                    if lvl == "DANGER":
                        st.markdown(f"<p style='color:#ef4444;font-weight:700;text-align:right;'>{chk.get('status_text', 'CHƯA ĐẠT')}</p>", unsafe_allow_html=True)
                    elif lvl == "WARNING":
                        st.markdown(f"<p style='color:#f59e0b;font-weight:700;text-align:right;'>{chk.get('status_text', 'LƯU Ý')}</p>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<p style='color:#22c55e;font-weight:700;text-align:right;'>{chk.get('status_text', 'ĐẠT CHUẨN')}</p>", unsafe_allow_html=True)

                st.write(chk['detail'])
                st.info(f"💡 **Hướng xử lý chuẩn:** {chk['recommendation']}")

        # ── 4. THÔNG TIN HỒ SƠ ĐÃ QUÉT ĐƯỢC ──
        with st.expander(f"📁 Danh sách {len(scanned_files)} tệp hồ sơ đã quét trong thư mục", expanded=False):
            if scanned_files:
                for f_path in scanned_files:
                    st.write(f"- `{os.path.basename(f_path)}` *({os.path.relpath(f_path, BASE_DIR)})*")
            else:
                st.info("Chưa có tệp scan nào trong thư mục công trình.")



# ── PAGE 3: BẢNG TỔNG HỢP QUYẾT TOÁN ──
elif "BẢNG TỔNG HỢP QUYẾT TOÁN" in page:
    st.markdown('<p class="page-title">📋 BẢNG TỔNG HỢP QUYẾT TOÁN KINH PHÍ SCL</p>', unsafe_allow_html=True)
    st.caption("Trích xuất và quản lý toàn diện 44 chỉ tiêu chi phí theo biểu mẫu chuẩn của EVNHCMC và Công ty Điện lực Vũng Tàu.")

    row_th, mr, ten, cd = _select_ct("global_project_sel")
    if row_th is not None:
        ma_ct = str(row_th.get('Mã CT', '')).strip()

        # Upload file khu vực
        with st.container(border=True):
            col_up, col_info = st.columns([2, 3])
            with col_up:
                uploaded_thqt = st.file_uploader(
                    "📤 Tải lên file Bảng THQT (.xlsx / .xls)",
                    type=["xlsx", "xls"],
                    key=f"uploader_thqt_{ma_ct}",
                    help="Chọn file Excel Bảng THQT kinh phí SCL đã điền số liệu quyết toán"
                )
            with col_info:
                st.markdown("**Hướng dẫn:**")
                st.markdown(
                    "- Khi công trình hoàn tất, anh tải file **Bảng THQT KP SCL (.xlsx)** lên tại đây.<br>"
                    "- Hệ thống sẽ **tự động bóc tách toàn bộ 44 danh mục chi phí**, tính toán các số liệu và **tự động điền đồng bộ vào 3 mẫu**:<br>"
                    "&nbsp;&nbsp;+ *Tờ trình duyệt quyết toán danh mục SCL*<br>"
                    "&nbsp;&nbsp;+ *Báo cáo thẩm tra quyết toán*<br>"
                    "&nbsp;&nbsp;+ *Quyết định phê duyệt quyết toán*",
                    unsafe_allow_html=True
                )

        if uploaded_thqt is not None:
            try:
                raw_bytes = uploaded_thqt.getvalue()
                parsed_res = parse_thqt_excel(raw_bytes)
                save_thqt_cache(ma_ct, parsed_res, raw_bytes)
                st.session_state[f'thqt_{ma_ct}'] = parsed_res
                st.toast("✅ Đã nạp và lưu dữ liệu Bảng THQT thành công!", icon="🎉")
            except Exception as e:
                st.error(f"Lỗi khi đọc file Excel: {e}")

        # Lấy dữ liệu từ session hoặc cache
        thqt_data = st.session_state.get(f'thqt_{ma_ct}')
        if not thqt_data:
            thqt_data = load_thqt_cache(ma_ct)
            if thqt_data:
                st.session_state[f'thqt_{ma_ct}'] = thqt_data

        if thqt_data and "summary" in thqt_data:
            sm = thqt_data["summary"]
            
            # 4 Thẻ chỉ số chính
            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(f'<div class="metric-card"><div class="metric-label">Chi phí Sửa chữa (Trước thuế)</div><div class="metric-val">{_fmt_money_dot(sm["sc_qt"])} đ</div><div style="font-size:11px;color:rgba(255,255,255,0.8);margin-top:4px;">Dự toán: {_fmt_money_dot(sm["sc_dt"])} đ</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="metric-card"><div class="metric-label">Chi phí KTCB khác</div><div class="metric-val">{_fmt_money_dot(sm["khac_qt"])} đ</div><div style="font-size:11px;color:rgba(255,255,255,0.8);margin-top:4px;">Dự toán: {_fmt_money_dot(sm["khac_dt"])} đ</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="metric-card"><div class="metric-label">Thuế GTGT</div><div class="metric-val">{_fmt_money_dot(sm["thue_qt"])} đ</div><div style="font-size:11px;color:rgba(255,255,255,0.8);margin-top:4px;">Dự toán: {_fmt_money_dot(sm["thue_dt"])} đ</div></div>', unsafe_allow_html=True)
            c4.markdown(f'<div class="metric-card"><div class="metric-label">Tổng giá trị Quyết toán</div><div class="metric-val">{_fmt_money_dot(sm["tong_qt"])} đ</div><div style="font-size:11px;color:rgba(255,255,255,0.8);margin-top:4px;">Bằng chữ: {sm["bang_chu"]}</div></div>', unsafe_allow_html=True)

            st.write("")
            c_sub1, c_sub2, c_sub3, c_sub4 = st.columns(4)
            c_sub1.info(f"⚙️ **Chi phí thiết bị, VT:** `{_fmt_money_dot(sm['tb_qt'])} đ` (DT: `{_fmt_money_dot(sm['tb_dt'])} đ`)")
            c_sub2.info(f"💵 **Tổng trước thuế:** `{_fmt_money_dot(sm['truoc_thue_qt'])} đ` (DT: `{_fmt_money_dot(sm['truoc_thue_dt'])} đ`)")
            c_sub3.warning(f"🔄 **Vật tư thu hồi:** `{_fmt_money_dot(sm['vtth_qt'])} đ` (DT: `{_fmt_money_dot(sm['vtth_dt'])} đ`)")
            c_sub4.success(f"⚡ **Chi phí SCL thực tế:** `{_fmt_money_dot(sm['scl_qt'])} đ` (DT: `{_fmt_money_dot(sm['scl_dt'])} đ`)")

            st.markdown('<p class="section-title">📊 Chi tiết toàn bộ 44 danh mục chi phí (Bảng THQT KP SCL)</p>', unsafe_allow_html=True)
            
            # Render HTML Table 44 rows
            rows_html = ""
            for item in thqt_data.get("rows_data", []):
                r_name = item["name"]
                is_bold = any(k in r_name for k in ["I.", "II.", "III.", "IV.", "V.", "VI.", "VII.", "VIII.", "TỔNG", "CỘNG", "Giá trị dự toán", "Chi phí khác"])
                style_row = ' class="row-header"' if is_bold else ""
                
                rows_html += f'''<tr{style_row}>
                    <td class="col-ten" style="font-weight:{'700' if is_bold else '400'};">{r_name}</td>
                    <td class="col-num">{_fmt_money_dot(item['dt']) if item['dt'] > 0 else '-'}</td>
                    <td class="col-num">{_fmt_money_dot(item['ps_tang']) if item['ps_tang'] > 0 else '-'}</td>
                    <td class="col-num">{_fmt_money_dot(item['ps_giam']) if item['ps_giam'] > 0 else '-'}</td>
                    <td class="col-num">{_fmt_money_dot(item['tong_dt']) if item['tong_dt'] > 0 else '-'}</td>
                    <td class="col-num" style="font-weight:{'700' if is_bold else '400'};color:#1565C0;">{_fmt_money_dot(item['qt']) if item['qt'] > 0 else ('0' if is_bold else '-')}</td>
                    <td style="font-size:11px;color:#64748b;">{item.get('note', '')}</td>
                </tr>\n'''

            table_thqt_html = f'''<div class="qt-table-wrapper" style="max-height:600px;overflow-x:auto;">
            <table class="qt-table" style="min-width:980px;">
                <thead><tr>
                    <th style="min-width:320px;">Nội dung chi phí</th>
                    <th style="min-width:110px;">Dự toán duyệt</th>
                    <th style="min-width:90px;">PS Tăng</th>
                    <th style="min-width:90px;">PS Giảm</th>
                    <th style="min-width:110px;">Tổng DT duyệt</th>
                    <th style="min-width:120px;">Quyết toán đề nghị</th>
                    <th style="min-width:100px;">Ghi chú</th>
                </tr></thead>
                <tbody>{rows_html}</tbody>
            </table></div>'''
            st.markdown(table_thqt_html, unsafe_allow_html=True)

            # Export Excel
            df_export = pd.DataFrame([{
                "Nội dung chi phí": it["name"],
                "Dự toán duyệt": it["dt"],
                "Phát sinh tăng": it["ps_tang"],
                "Phát sinh giảm": it["ps_giam"],
                "Tổng dự toán duyệt": it["tong_dt"],
                "Quyết toán đề nghị": it["qt"],
                "Ghi chú": it["note"]
            } for it in thqt_data.get("rows_data", [])])
            
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Bang_THQT')
            safe_name = clean_filename(ten)
            st.markdown(create_download_link(buf.getvalue(), f"Bang_THQT_{safe_name}.xlsx", "📥 Xuất Excel - Bảng THQT KP SCL", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), unsafe_allow_html=True)
            
            st.success("🔗 **Đã liên kết dữ liệu Bảng THQT**: Khi anh chuyển sang tab **'📝 TỜ TRÌNH DUYỆT QT'**, **'📜 BÁO CÁO THẨM TRA'**, hoặc **'⚖️ QUYẾT ĐỊNH PHÊ DUYỆT'**, toàn bộ số liệu trên đã được điền tự động 100%!")
        else:
            st.info("💡 Chưa có dữ liệu Bảng THQT cho công trình này. Anh hãy bấm nút **'Browse files'** ở trên để tải file Excel **Bảng THQT KP SCL** vào nhé.")
            sample_file = os.path.join(BASE_DIR, "HS_mau", "2. Bảng THQT KP SCL.xlsx")
            if os.path.exists(sample_file):
                if st.button("📂 Nạp nhanh file Bảng THQT mẫu (HS_mau)", key=f"btn_quick_thqt_{ma_ct}"):
                    res_sample = parse_thqt_excel(sample_file)
                    save_thqt_cache(ma_ct, res_sample)
                    st.session_state[f'thqt_{ma_ct}'] = res_sample
                    st.rerun()

# ── PAGE 4: TỜ TRÌNH DUYỆT QUYẾT TOÁN DANH MỤC SCL ──
elif "TỜ TRÌNH" in page:
    st.markdown('<p class="page-title">📝 TỜ TRÌNH DUYỆT QUYẾT TOÁN DANH MỤC SCL HOÀN THÀNH</p>', unsafe_allow_html=True)
    row_th, mr, ten, cd = _select_ct("global_project_sel")
    if row_th is not None:
        ma_ct = str(row_th.get('Mã CT', '')).strip()
        so_dt = str(mr.get('Số Dự toán', '135/QĐ-PCVT')) if mr is not None and pd.notna(mr.get('Số Dự toán')) else '135/QĐ-PCVT'
        ngay_dt = _format_date_vn(mr.get('Ngày Dự toán')) if mr is not None else '16/03/2026'
        if '....' in ngay_dt: ngay_dt = '16/03/2026'
        now = datetime.datetime.now()

        # Nạp dữ liệu Bảng THQT
        thqt_data = st.session_state.get(f'thqt_{ma_ct}') or load_thqt_cache(ma_ct)
        if thqt_data and "summary" in thqt_data:
            sm = thqt_data["summary"]
            if thqt_data.get('ten_ct'): ten = thqt_data['ten_ct']
            if thqt_data.get('ma_ct'): ma_ct = thqt_data['ma_ct']
            sc_qt = sm.get("sc_qt", 0)
            tb_qt = sm.get("tb_qt", 0)
            khac_qt = sm.get("khac_qt", 0)
            truoc_thue = sm.get("truoc_thue_qt", 0)
            thue_qt = sm.get("thue_qt", 0)
            tong_qt = sm.get("tong_qt", 0)
            vtth_qt = sm.get("vtth_qt", 0)
            scl_qt = sm.get("scl_qt", 0)
            bang_chu = sm.get("bang_chu") or doc_so_vn(tong_qt)
            is_from_thqt = True
        else:
            bd = get_cost_breakdown(cd)
            scl_qt = bd.get('SCL', {}).get('qt', 0)
            if scl_qt == 0 and mr is not None and 'Giá trị Q.định phê duyệt QT công trình' in mr:
                scl_qt = float(mr.get('Giá trị Q.định phê duyệt QT công trình', 0))
            thue_qt = bd.get('B.8', {}).get('qt', 0)
            sc_qt = bd.get('B', {}).get('qt', 0)
            tb_qt = bd.get('A', {}).get('qt', 0)
            khac_qt = bd.get('C', {}).get('qt', 0)
            vtth_qt = bd.get('F', {}).get('qt', 0)
            if thue_qt == 0 and scl_qt > 0: thue_qt = int(round(scl_qt - scl_qt / 1.08))
            truoc_thue = scl_qt - thue_qt if scl_qt >= thue_qt else scl_qt
            tong_qt = scl_qt
            bang_chu = doc_so_vn(tong_qt)
            is_from_thqt = False

        if is_from_thqt:
            st.success(f"⚡ Số liệu đã được **tự động điền 100% từ Bảng THQT KP SCL** của công trình `{ma_ct}`!")
        else:
            st.info("💡 Chưa tải lên file Bảng THQT cho công trình này. Hệ thống đang lấy số liệu tạm từ cơ sở dữ liệu.")

        st.markdown('<p class="section-title">Xem trước nội dung Tờ trình duyệt quyết toán (Chuẩn HS mẫu)</p>', unsafe_allow_html=True)
        preview_html = f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:15pt;">
                <tr>
                    <td style="width:45%;text-align:center;border:none;vertical-align:top;font-size:11pt;">
                        CÔNG TY ĐIỆN LỰC VŨNG TÀU<br>
                        <b>PHÒNG TÀI CHÍNH KẾ TOÁN</b><br>
                        -----------------------<br><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /TTr-TCKT<br>
                        <i>V/v đề nghị duyệt quyết toán danh<br>mục SCL hoàn thành</i>
                    </td>
                    <td style="width:55%;text-align:center;border:none;vertical-align:top;">
                        <b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b>Độc lập – Tự do – Hạnh phúc</b><br>
                        -----------------------<br><br>
                        <i>Vũng Tàu, ngày &nbsp;&nbsp;&nbsp; tháng &nbsp;&nbsp;&nbsp; năm {now.year}</i>
                    </td>
                </tr>
            </table>
            
            <p style="margin-top:20pt;font-weight:bold;">Kính gửi : Giám đốc Công ty Điện lực Vũng Tàu</p>
            
            <p style="text-align:justify;line-height:1.6;margin-top:15pt;">
            Căn cứ Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Tổng công ty Điện lực TP.HCM về việc ban hành quy định thực hiện công tác sửa chữa lớn tài sản trong Tổng công ty Điện lực Thành phố Hồ Chí Minh;<br><br>
            Căn cứ  quyết định số {so_dt} ngày {ngay_dt} về việc phê duyệt bổ sung danh mục công trình và điều hòa kế hoạch vốn sửa chữa lớn năm {now.year} của Công ty Điện lực Vũng Tàu;<br><br>
            Căn cứ báo cáo thẩm tra quyết toán công trình {ten}. Kính đề nghị Ông Giám đốc phê duyệt quyết toán công trình {ten}, mã công trình {ma_ct} với tổng giá trị công trình là: <b>{_fmt_money_dot(tong_qt)} đồng.</b><br>
            (Bằng chữ:  {bang_chu}).<br>
            Trong đó:<br>
            </p>
            {_render_cost_table_html([
                ("Chi phí Sửa chữa", _fmt_money_dot(sc_qt), False, "đồng."),
                ("Chi phí vật tư, thiết bị", _fmt_money_dot(tb_qt), False, "đồng."),
                ("Chi phí KTCB khác", _fmt_money_dot(khac_qt), False, "đồng."),
                ("Tổng giá trị trước thuế", _fmt_money_dot(truoc_thue), True, "đồng."),
                ("Thuế GTGT", _fmt_money_dot(thue_qt), False, "đồng."),
                ("Tổng giá trị Quyết toán", _fmt_money_dot(tong_qt), True, "đồng."),
                ("Chi phí vật tư thu hồi", _fmt_money_dot(vtth_qt), False, "đồng."),
                ("Chi phí sửa chữa lớn", _fmt_money_dot(scl_qt), True, "đồng"),
            ], indent_px=35, label_min_width=210)}
            <p style="text-align:justify;line-height:1.6;margin-top:8pt;">
            Trân trọng./.
            </p>
            
            <table style="width:100%;border:none;margin-top:30pt;">
                <tr>
                    <td style="width:50%;border:none;vertical-align:top;font-size:10pt;">
                        <b>Nơi nhận:</b><br>
                        - Như trên<br>
                        - Lưu: VT, TCKT. HMT.
                    </td>
                    <td style="width:50%;text-align:center;border:none;vertical-align:top;">
                        <b>TRƯỞNG PHÒNG TÀI CHÍNH KẾ TOÁN</b><br><br><br><br><br>
                        <b>Hà Thị Mai Hiên</b>
                    </td>
                </tr>
            </table>
        </div>
        '''
        data_ttr = export_ttr_duyet_qt_word(mr, cd, thqt_data)
        safe = clean_filename(ten)
        st.markdown(create_download_link(data_ttr, f"TTr_Duyet_QT_{safe}.docx", "📥 Xuất Word - Tờ trình duyệt QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
        st.write('')
        components.html(wrap_preview_html(preview_html), height=750, scrolling=True)

# ── PAGE 5: BÁO CÁO THẨM TRA ──
elif "BÁO CÁO THẨM TRA" in page or "BÁO CÁO THẨN TRA" in page or "THẨM TRA" in page:
    st.markdown('<p class="page-title">📜 BÁO CÁO THẨM TRA QUYẾT TOÁN CÔNG TRÌNH SCL</p>', unsafe_allow_html=True)
    row_th, mr, ten, cd = _select_ct("global_project_sel")
    if row_th is not None:
        ma_ct = str(row_th.get('Mã CT', '')).strip()
        now = datetime.datetime.now()

        # Nạp dữ liệu Bảng THQT
        thqt_data = st.session_state.get(f'thqt_{ma_ct}') or load_thqt_cache(ma_ct)
        if thqt_data and "summary" in thqt_data:
            sm = thqt_data["summary"]
            if thqt_data.get('ten_ct'): ten = thqt_data['ten_ct']
            if thqt_data.get('ma_ct'): ma_ct = thqt_data['ma_ct']
            sc_dt = sm.get("sc_dt", 0)
            sc_qt = sm.get("sc_qt", 0)
            tb_dt = sm.get("tb_dt", 0)
            tb_qt = sm.get("tb_qt", 0)
            khac_dt = sm.get("khac_dt", 0)
            khac_qt = sm.get("khac_qt", 0)
            truoc_thue = sm.get("truoc_thue_qt", 0)
            thue_qt = sm.get("thue_qt", 0)
            tong_qt = sm.get("tong_qt", 0)
            vtth_qt = sm.get("vtth_qt", 0)
            scl_qt = sm.get("scl_qt", 0)
            ktcb_list = thqt_data.get("ktcb_rows", [])
            tot_ktcb_dt = sm.get("tot_ktcb_dt", 0)
            tot_ktcb_qt = sm.get("tot_ktcb_qt", 0)
            is_from_thqt = True
        else:
            bd = get_cost_breakdown(cd)
            scl_dt = bd.get('SCL', {}).get('dt', 0)
            scl_qt = bd.get('SCL', {}).get('qt', 0)
            if scl_dt == 0 and mr is not None and 'Giá trị Dự toán' in mr:
                scl_dt = float(mr.get('Giá trị Dự toán', 0))
            if scl_qt == 0 and mr is not None and 'Giá trị Q.định phê duyệt QT công trình' in mr:
                scl_qt = float(mr.get('Giá trị Q.định phê duyệt QT công trình', 0))
            sc_dt = bd.get('B', {}).get('dt', 0)
            sc_qt = bd.get('B', {}).get('qt', 0)
            tb_dt = bd.get('A', {}).get('dt', 0)
            tb_qt = bd.get('A', {}).get('qt', 0)
            khac_dt = bd.get('C', {}).get('dt', 0)
            khac_qt = bd.get('C', {}).get('qt', 0)
            vtth_qt = bd.get('F', {}).get('qt', 0)
            thue_qt = bd.get('B.8', {}).get('qt', 0)
            if thue_qt == 0 and scl_qt > 0: thue_qt = int(round(scl_qt - scl_qt / 1.08))
            truoc_thue = scl_qt - thue_qt if scl_qt >= thue_qt else scl_qt
            tong_qt = scl_qt
            ktcb_list = []
            tot_ktcb_dt = khac_dt
            tot_ktcb_qt = khac_qt
            is_from_thqt = False

        if is_from_thqt:
            st.success(f"⚡ Số liệu đã được **tự động điền 100% từ Bảng THQT KP SCL** của công trình `{ma_ct}`!")
        else:
            st.info("💡 Chưa tải lên file Bảng THQT cho công trình này. Hệ thống đang lấy số liệu tạm từ cơ sở dữ liệu.")

        # Xây dựng bảng 5 cột KTCB khác
        if not ktcb_list:
            ktcb_list = [
                {"name": "Chi phí Lập khảo sát, Lập PAKT-DT", "dt": 46296296, "qt": 46296296, "tt": 46296296, "diff": 0},
                {"name": "2- Chi phí giám sát công trình", "dt": 38888889, "qt": 38888889, "tt": 38888889, "diff": 0},
                {"name": "Chi phí bảo hiểm công trình", "dt": 898673, "qt": 0, "tt": 0, "diff": 0},
                {"name": "Chi phí thẩm tra thiết kế", "dt": 2107797, "qt": 0, "tt": 0, "diff": 0},
                {"name": "Chi phí thẩm tra dự toán", "dt": 2042439, "qt": 0, "tt": 0, "diff": 0},
                {"name": "Chi phí thẩm tra phê duyệt quyết toán", "dt": 5619375, "qt": 0, "tt": 0, "diff": 0},
                {"name": "Chi phí dự phòng", "dt": 49292765, "qt": 0, "tt": 0, "diff": 0},
            ]
            tot_ktcb_dt = sum(x["dt"] for x in ktcb_list)
            tot_ktcb_qt = sum(x["qt"] for x in ktcb_list)

        ktcb_rows_html = ""
        for it in ktcb_list:
            ktcb_rows_html += f'''<tr>
                <td style="border:1px solid #000;padding:4pt 6pt;">{it['name']}</td>
                <td style="border:1px solid #000;padding:4pt 6pt;text-align:right;">{_fmt_money_dot(it.get('dt',0))}</td>
                <td style="border:1px solid #000;padding:4pt 6pt;text-align:right;">{_fmt_money_dot(it.get('qt',0))}</td>
                <td style="border:1px solid #000;padding:4pt 6pt;text-align:right;">{_fmt_money_dot(it.get('tt', it.get('qt',0)))}</td>
                <td style="border:1px solid #000;padding:4pt 6pt;text-align:right;">0</td>
            </tr>\n'''
        ktcb_rows_html += f'''<tr style="font-weight:bold;background:#f0f4f8;">
            <td style="border:1px solid #000;padding:4pt 6pt;">Tổng cộng</td>
            <td style="border:1px solid #000;padding:4pt 6pt;text-align:right;">{_fmt_money_dot(tot_ktcb_dt)}</td>
            <td style="border:1px solid #000;padding:4pt 6pt;text-align:right;">{_fmt_money_dot(tot_ktcb_qt)}</td>
            <td style="border:1px solid #000;padding:4pt 6pt;text-align:right;">{_fmt_money_dot(tot_ktcb_qt)}</td>
            <td style="border:1px solid #000;padding:4pt 6pt;text-align:right;">0</td>
        </tr>\n'''

        st.markdown('<p class="section-title">Xem trước Báo cáo thẩm tra quyết toán (Chuẩn HS mẫu)</p>', unsafe_allow_html=True)
        preview5a = f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:10pt;">
                <tr>
                    <td style="width:45%;text-align:center;border:none;vertical-align:top;font-size:10pt;">
                        TỔNG CÔNG TY ĐIỆN LỰC TP HỒ CHÍ MINH<br>
                        <b>CÔNG TY ĐIỆN LỰC VŨNG TÀU</b><br><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /BC-TCKT
                    </td>
                    <td style="width:55%;text-align:center;border:none;vertical-align:top;">
                        <b>CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b>Độc lập - Tự do - Hạnh phúc</b><br>
                        -----------------------<br><br>
                        <i>Vũng Tàu, ngày &nbsp;&nbsp;&nbsp; tháng &nbsp;&nbsp;&nbsp; năm {now.year}.</i>
                    </td>
                </tr>
            </table>

            <div class="center-title" style="margin-top:15pt;font-weight:bold;font-size:14pt;text-align:center;">
                BÁO CÁO THẨM TRA QUYẾT TOÁN<br>CÔNG TRÌNH SỬA CHỮA LỚN
            </div>
            
            <p>Căn cứ hồ sơ quyết toán công trình: {ten}.</p>
            <p>Tổ thẩm tra quyết toán công trình sửa chữa lớn Công ty Điện lực Vũng Tàu thẩm tra hồ sơ quyết toán công trình với kết quả như sau:</p>
            
            <p><b><i>I/- Nội dung thẩm tra:</i></b></p>
            <p style="margin-left:20px;">Tên công trình: {ten}.</p>
            <p style="margin-left:20px;">Đơn vị quản lý: Công ty Điện lực Vũng Tàu.</p>
            <p style="margin-left:20px;">Phương thức chọn thầu thực hiện: đấu thầu rộng rãi</p>
            <p style="margin-left:20px;">Hợp đồng số:<br>40-2026/HĐXL/PCVT-TB.<br>41-2026/HĐTV/PCVT-TVĐ.<br>03-2026/HĐTV/PCVT-TVĐ.</p>
            <p style="margin-left:20px;">Đơn vị thực hiện:<br>Công ty CP Tư vấn Dịch vụ Xây dựng Thái Bình.<br>Công ty CP Tư vấn Xây dựng Điện lực TP HCM.</p>
            
            <p><b>Nội dung cụ thể:</b></p>
            <p style="margin-left:20px;margin-bottom:2px;"><b>1/ Phần sửa chữa:</b></p>
            {_render_cost_table_html([
                ("- Số dự toán", _fmt_money_dot(sc_dt), False, "đồng."),
                ("- Số quyết toán", _fmt_money_dot(sc_qt), False, "đồng."),
                ("- Số thẩm tra", _fmt_money_dot(sc_qt), False, "đồng."),
                ("- Số chênh lệch TTr/QT", "0", False, "đồng."),
            ], indent_px=35, label_min_width=175)}
            
            <p style="margin-left:20px;margin-bottom:2px;margin-top:6pt;"><b>2/  Phần vật tư, thiết bị:</b></p>
            {_render_cost_table_html([
                ("- Số dự toán", _fmt_money_dot(tb_dt), False, "đồng."),
                ("- Số quyết toán", _fmt_money_dot(tb_qt), False, "đồng."),
                ("- Số thẩm tra", _fmt_money_dot(tb_qt), False, "đồng."),
                ("- Số chênh lệch TTr/QT", "0", False, "đồng."),
            ], indent_px=35, label_min_width=175)}
            
            <p style="margin-left:20px;margin-top:6pt;"><b>3/ Phần kiến thiết cơ bản khác:</b></p>
            <table style="width:100%;border-collapse:collapse;margin:10pt 0;font-size:11pt;">
                <tr style="font-weight:bold;text-align:center;background:#f0f0f0;">
                    <td style="border:1px solid #000;padding:5pt;">Nội dung</td>
                    <td style="border:1px solid #000;padding:5pt;width:18%;">Dự toán được duyệt</td>
                    <td style="border:1px solid #000;padding:5pt;width:18%;">Giá trị quyết toán</td>
                    <td style="border:1px solid #000;padding:5pt;width:18%;">Số thẩm tra</td>
                    <td style="border:1px solid #000;padding:5pt;width:18%;">Chênh lệch TTr/QT</td>
                </tr>
                {ktcb_rows_html}
            </table>

            <p><b><i>II/ Kết luận:</i></b></p>
            <p style="margin-left:20px;margin-bottom:4pt;">Sau khi xem xét thẩm tra hồ sơ, tổ thẩm tra quyết toán chấp thuận tổng giá trị quyết toán công trình nêu trên là: <b>{_fmt_money_dot(tong_qt)} đồng</b>, cụ thể:</p>
            {_render_cost_table_html([
                ("Chi phí Sửa chữa", _fmt_money_dot(sc_qt), False, "đồng."),
                ("Chi phí vật tư, thiết bị", _fmt_money_dot(tb_qt), False, "đồng."),
                ("Chi phí KTCB khác", _fmt_money_dot(khac_qt), False, "đồng."),
                ("Tổng giá trị trước thuế", _fmt_money_dot(truoc_thue), True, "đồng."),
                ("Thuế GTGT", _fmt_money_dot(thue_qt), False, "đồng."),
                ("Tổng giá trị Quyết toán", _fmt_money_dot(tong_qt), True, "đồng."),
                ("Chi phí vật tư thu hồi", _fmt_money_dot(vtth_qt), False, "đồng."),
                ("Chi phí sửa chữa lớn", _fmt_money_dot(scl_qt), True, "đồng"),
            ], indent_px=40, label_min_width=210)}
            <p style="margin-left:20px;margin-top:10pt;">Kính trình và đề nghị Ông Giám đốc phê duyệt.</p>

            <div style="text-align:center;font-weight:bold;margin:25pt 0 15pt 0;">THÀNH VIÊN TỔ THẨM TRA QUYẾT TOÁN</div>
            <table style="width:100%;border:none;">
                <tr>
                    <td style="width:60%;border:none;line-height:2.0;font-size:11pt;">
                        Hà Thị Mai Hiên (KTT)…………………….…...........<br>
                        Nguyễn văn Quyến (TP.QLĐT)………………………<br>
                        Nguyễn Mạnh Hiệp (TP.KHVT)……………………...<br>
                        Nguyễn Đức Minh (TP.KTAT)……………………….<br>
                        Hoàng Minh Tuấn (CV.TCKT)…………….…………<br>
                        Nguyễn Văn Huy (TT.KTAT)………………………<br>
                        Lê Anh Ngọc (TT.KHVT)…………………………….<br>
                        Lê Hoài Nam (CV.QLĐT)………………………
                    </td>
                    <td style="width:40%;text-align:center;border:none;vertical-align:top;">
                        <b>TỔ TRƯỞNG TỔ THẨM TRA</b><br><br><br><br><br><br><br><br>
                        <b>Trần Thanh Hải</b>
                    </td>
                </tr>
            </table>
        </div>
        '''
        data_bc = export_bao_cao_tham_tra_word(mr, cd, thqt_data)
        safe = clean_filename(ten)
        st.markdown(create_download_link(data_bc, f"BC_tham_tra_{safe}.docx", "📥 Xuất Word - Báo cáo thẩm tra", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
        st.write('')
        components.html(wrap_preview_html(preview5a), height=950, scrolling=True)

# ── PAGE 6: QUYẾT ĐỊNH PHÊ DUYỆT QUYẾT TOÁN ──
elif "QUYẾT ĐỊNH" in page:
    st.markdown('<p class="page-title">⚖️ QUYẾT ĐỊNH PHÊ DUYỆT QUYẾT TOÁN CÔNG TRÌNH SCL</p>', unsafe_allow_html=True)
    row_th, mr, ten, cd = _select_ct("global_project_sel")
    if row_th is not None:
        ma_ct = str(row_th.get('Mã CT', '')).strip()
        so_dt = str(mr.get('Số Dự toán', '135/QĐ-PCVT')) if mr is not None and pd.notna(mr.get('Số Dự toán')) else '135/QĐ-PCVT'
        ngay_dt = _format_date_vn(mr.get('Ngày Dự toán')) if mr is not None else '16/03/2026'
        if '....' in ngay_dt: ngay_dt = '16/03/2026'
        now = datetime.datetime.now()

        # Nạp dữ liệu Bảng THQT
        thqt_data = st.session_state.get(f'thqt_{ma_ct}') or load_thqt_cache(ma_ct)
        if thqt_data and "summary" in thqt_data:
            sm = thqt_data["summary"]
            if thqt_data.get('ten_ct'): ten = thqt_data['ten_ct']
            if thqt_data.get('ma_ct'): ma_ct = thqt_data['ma_ct']
            sc_qt = sm.get("sc_qt", 0)
            tb_qt = sm.get("tb_qt", 0)
            khac_qt = sm.get("khac_qt", 0)
            truoc_thue = sm.get("truoc_thue_qt", 0)
            thue_qt = sm.get("thue_qt", 0)
            tong_qt = sm.get("tong_qt", 0)
            vtth_qt = sm.get("vtth_qt", 0)
            scl_qt = sm.get("scl_qt", 0)
            bang_chu = sm.get("bang_chu") or doc_so_vn(tong_qt)
            is_from_thqt = True
        else:
            bd = get_cost_breakdown(cd)
            scl_qt = bd.get('SCL', {}).get('qt', 0)
            if scl_qt == 0 and mr is not None and 'Giá trị Q.định phê duyệt QT công trình' in mr:
                scl_qt = float(mr.get('Giá trị Q.định phê duyệt QT công trình', 0))
            tb_qt = bd.get('A', {}).get('qt', 0)
            sc_qt = bd.get('B', {}).get('qt', 0)
            khac_qt = bd.get('C', {}).get('qt', 0)
            vtth_qt = bd.get('F', {}).get('qt', 0)
            thue_qt = bd.get('B.8', {}).get('qt', 0)
            if thue_qt == 0 and scl_qt > 0: thue_qt = int(round(scl_qt - scl_qt / 1.08))
            truoc_thue = scl_qt - thue_qt if scl_qt >= thue_qt else scl_qt
            tong_qt = scl_qt
            bang_chu = doc_so_vn(scl_qt) if scl_qt > 0 else ''
            is_from_thqt = False

        if is_from_thqt:
            st.success(f"⚡ Số liệu đã được **tự động điền 100% từ Bảng THQT KP SCL** của công trình `{ma_ct}`!")
        else:
            st.info("💡 Chưa tải lên file Bảng THQT cho công trình này. Hệ thống đang lấy số liệu tạm từ cơ sở dữ liệu.")

        st.markdown('<p class="section-title">Xem trước Quyết định phê duyệt quyết toán (Chuẩn HS mẫu)</p>', unsafe_allow_html=True)
        preview5b = f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:10pt;">
                <tr>
                    <td style="width:45%;text-align:center;border:none;vertical-align:top;font-size:11pt;">
                        TỔNG CÔNG TY<br>
                        ĐIỆN LỰC TP HỒ CHÍ MINH<br>
                        <b>CÔNG TY ĐIỆN LỰC VŨNG TÀU</b><br><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /QĐ-PCVT
                    </td>
                    <td style="width:55%;text-align:center;border:none;vertical-align:top;">
                        <b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b>Độc lập – Tự do – Hạnh phúc</b><br>
                        -----------------------<br><br>
                        <i>Vũng Tàu, ngày &nbsp;&nbsp;&nbsp; tháng &nbsp;&nbsp;&nbsp; năm {now.year}</i>
                    </td>
                </tr>
            </table>

            <div class="center-title" style="margin-top:15pt;font-weight:bold;font-size:14pt;text-align:center;">
                QUYẾT ĐỊNH<br>
                <span style="font-size:12pt;font-weight:normal;">V/v phê duyệt quyết toán công trình Sửa chữa lớn hoàn thành</span>
            </div>
            
            <p>Tên công trình: <b>{ten}</b></p>
            <p>Mã công trình: <b>{ma_ct}</b></p>
            <p>Chủ đầu tư: <b>Công ty Điện lực Vũng Tàu</b></p>
            <p>Kế hoạch vốn năm: <b>{now.year}</b></p>
            <p>Nguồn vốn: <b>Sửa chữa lớn</b></p>
            
            <div style="text-align:center;font-weight:bold;margin:15pt 0 10pt 0;">GIÁM ĐỐC CÔNG TY ĐIỆN LỰC VŨNG TÀU</div>
            
            <p style="text-align:justify;line-height:1.6;">
            Căn cứ Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Tổng công ty Điện lực TP.HCM về việc ban hành quy định thực hiện công tác sửa chữa lớn tài sản trong Tổng công ty Điện lực Thành phố Hồ Chí Minh;<br>
            Căn cứ  quyết định số {so_dt} ngày {ngay_dt} về việc phê duyệt bổ sung danh mục công trình và điều hòa kế hoạch vốn sửa chữa lớn năm {now.year} của Công ty Điện lực Vũng Tàu;<br>
            Căn cứ hồ sơ quyết toán công trình: {ten};<br>
            Căn cứ báo cáo thẩm tra quyết toán của Tổ thẩm tra duyệt quyết toán công trình sửa chữa lớn Công ty Điện lực Vũng tàu.
            </p>
            
            <div style="text-align:center;font-weight:bold;margin:15pt 0 10pt 0;">QUYẾT ĐỊNH :</div>
            
            <p><b>Điều 1. Phê duyệt quyết toán</b></p>
            <p>Công trình: <b>{ten}</b> với tổng giá trị công trình : <b>{_fmt_money_dot(tong_qt)} đồng.</b></p>
            <p>(Bằng chữ: <i>{bang_chu}</i>).</p>
            <p style="margin-left:20px;margin-bottom:2px;">Trong đó:</p>
            {_render_cost_table_html([
                ("Chi phí Sửa chữa", _fmt_money_dot(sc_qt), False, "đồng."),
                ("Chi phí vật tư, thiết bị", _fmt_money_dot(tb_qt), False, "đồng."),
                ("Chi phí KTCB khác", _fmt_money_dot(khac_qt), False, "đồng."),
                ("Tổng giá trị trước thuế", _fmt_money_dot(truoc_thue), True, "đồng."),
                ("Thuế GTGT", _fmt_money_dot(thue_qt), False, "đồng."),
                ("Tổng giá trị Quyết toán", _fmt_money_dot(tong_qt), True, "đồng."),
                ("Chi phí vật tư thu hồi", _fmt_money_dot(vtth_qt), False, "đồng."),
                ("Chi phí sửa chữa lớn", _fmt_money_dot(scl_qt), True, "đồng"),
            ], indent_px=35, label_min_width=210)}
            
            <p><b>Điều 2. Nguồn vốn thực hiện công trình</b></p>
            <p>Nguồn vốn: Sửa chữa lớn năm {now.year} của Công ty Điện lực Vũng Tàu.</p>
            
            <p><b>Điều 3. Thực hiện</b></p>
            <p>Phòng Tài chính Kế toán, Phòng Kế hoạch Vật tư, Phòng Quản lý đầu tư, Phòng Kỹ Thuật An Toàn Công ty Điện lực Vũng Tàu chịu trách nhiệm thi hành quyết định này./.</p>
            
            <table style="width:100%;border:none;margin-top:30pt;">
                <tr>
                    <td style="width:50%;border:none;font-size:10pt;vertical-align:top;">
                        <b>Nơi nhận:</b><br>
                        - P.KHVT, P.QLĐT, P.KTAT (để thực hiện);<br>
                        - Lưu: VT, TCKT, HMT. (04)
                    </td>
                    <td style="width:50%;text-align:center;border:none;vertical-align:top;">
                        <b>GIÁM ĐỐC</b><br><br><br><br><br>
                        <b>Nguyễn Ngọc Tuyến</b>
                    </td>
                </tr>
            </table>
        </div>
        '''
        data_qd = export_qd_phe_duyet_word(mr, cd, thqt_data)
        safe = clean_filename(ten)
        st.markdown(create_download_link(data_qd, f"QD_phe_duyet_{safe}.docx", "📥 Xuất Word - QĐ phê duyệt QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
        st.write('')
        components.html(wrap_preview_html(preview5b), height=900, scrolling=True)

