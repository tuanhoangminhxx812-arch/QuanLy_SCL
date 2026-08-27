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
[data-testid="stSidebar"]{background:linear-gradient(180deg,#1565C0 0%,#1976D2 40%,#2196F3 100%)}
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3,
[data-testid="stSidebar"] .stMarkdown h5,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stMarkdown span,
[data-testid="stSidebar"] label {color:#ffffff!important}
[data-testid="stSidebar"] div[role="radiogroup"] {gap: 8px;}
[data-testid="stSidebar"] div[role="radiogroup"] > label {
    background: rgba(255, 255, 255, 0.08);
    border: 1px solid rgba(255, 255, 255, 0.15);
    border-radius: 10px;
    padding: 12px 15px;
    cursor: pointer;
    transition: all 0.2s ease;
    display: flex;
    align-items: center;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
    background: rgba(255, 255, 255, 0.2);
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}
[data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"] {
    background: #ffffff;
    border-left: 5px solid #ff9800;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"] p {
    color: #1565C0 !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] div[data-testid="stMarkdownContainer"] {
    margin-left: 5px;
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
</style>
"""
def wrap_preview_html(body_html):
    return f"<!DOCTYPE html><html><head><meta charset='utf-8'>{A4_CSS}</head><body>{body_html}</body></html>"

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
        ["📊 Tổng quan","⚖️ Kiểm dò pháp lý & Thanh toán","📋 Bảng tổng hợp QT","📝 Tờ trình duyệt QT","📄 Thuyết minh QT",
         "🔍 Phiếu thẩm tra","📜 Báo cáo thẩm tra","📜 Quyết định phê duyệt"],
        label_visibility="collapsed")
    st.divider()
    st.caption(f"Cập nhật: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}")



# ── Chọn CT helper ──
def _select_ct(key, container=None):
    if container is None: container = st
    names=[]
    for _,r in df_th.iterrows():
        ma=str(r.get('Mã CT','')).strip();ten=str(r.get('Tên công trình','')).strip()
        names.append(f"{ma} - {ten}")
    sel=container.selectbox("Chọn công trình:",names,key=key)
    if not sel:return None,None,None,pd.DataFrame()
    ma=sel.split(" - ")[0].strip()
    ri=df_th[df_th['Mã CT']==ma].index
    if len(ri)==0:return None,None,None,pd.DataFrame()
    row_th=df_th.loc[ri[0]]
    ten=str(row_th.get('Tên công trình',''))
    mr,cd=get_project_section(db_df,ten,ma) if not db_df.empty else (None,pd.DataFrame())
    return row_th,mr,ten,cd

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
header_container = st.container()

with header_container:
    st.markdown('<div class="header-content-wrapper"></div>', unsafe_allow_html=True)
    st.markdown('<h1 class="main-title">HỆ THỐNG QUẢN LÝ QUYẾT TOÁN SỬA CHỮA LỚN</h1>', unsafe_allow_html=True)

# ── PAGE 1: TỔNG QUAN ──
if page=="📊 Tổng quan":
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

        with header_container:
            c1,c2,c3,c4=st.columns(4)
            for col,lbl,val in [(c1,"TỔNG SỐ CÔNG TRÌNH",str(total_ct)),
                (c2,"TỔNG KHÁI TOÁN",fmt_money(total_kh)),
                (c3,"TỔNG THỰC HIỆN",fmt_money(total_th)),
                (c4,"TỶ LỆ GIẢI NGÂN",f"{ty_le:.1f}%")]:
                col.markdown(f'<div class="metric-card"><div class="metric-label">{lbl}</div><div class="metric-val">{val}</div></div>',unsafe_allow_html=True)

        # ── CẬP NHẬT DỮ LIỆU NHANH ──
        with st.expander("📤 Cập nhật dữ liệu báo cáo (CapNhatTienDo.xlsx / PM_092.xlsx)", expanded=False):
            up_col1, up_col2 = st.columns(2)
            with up_col1:
                st.markdown('**📋 File Tiến độ** (`CapNhatTienDo.xlsx`)')
                up_tiendo = st.file_uploader(
                    "Chọn file CapNhatTienDo.xlsx mới",
                    type=['xlsx','xls'],
                    key='up_capnhat_td',
                    label_visibility='collapsed'
                )
                if up_tiendo:
                    with st.spinner("Đang cập nhật CapNhatTienDo.xlsx..."):
                        ok = gh_upload_root_file('CapNhatTienDo.xlsx', up_tiendo.getvalue())
                    if ok:
                        st.success("✅ Đã cập nhật CapNhatTienDo.xlsx thành công!")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("❌ Lỗi khi cập nhật file.")
            with up_col2:
                st.markdown('**💰 File PM_092** (`PM_092.xlsx`)')
                up_pm092 = st.file_uploader(
                    "Chọn file PM_092.xlsx mới",
                    type=['xlsx','xls'],
                    key='up_pm092',
                    label_visibility='collapsed'
                )
                if up_pm092:
                    with st.spinner("Đang cập nhật PM_092.xlsx..."):
                        ok = gh_upload_root_file('PM_092.xlsx', up_pm092.getvalue())
                    if ok:
                        st.success("✅ Đã cập nhật PM_092.xlsx thành công!")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("❌ Lỗi khi cập nhật file.")

        # Health analysis
        now=datetime.datetime.now();cy,cm=now.year,now.month
        hd_list=[]
        for _,r in df_th.iterrows():
            hs,hi,hins,gn_rate=analyze_health(r,cy,cm)
            hd_list.append({'Mã CT':r.get('Mã CT',''),'Tên công trình':r.get('Tên công trình',''),
                'Sức khỏe':f"{hi} {hs}",'Đánh giá':hins, 'Tỷ lệ GN': gn_rate, '_s':hs})
        df_h=pd.DataFrame(hd_list)
        if 'Tên công trình' in df_h.columns:
            df_h['Tên công trình'] = df_h['Tên công trình'].apply(shorten_name)

        st.markdown('<p class="section-title">🚨 Đánh giá & Cảnh báo rủi ro</p>',unsafe_allow_html=True)
        with st.container(border=True):
            r1,r2,r3,r4=st.columns(4)
            r1.metric("🟢 Hoàn thành",len(df_h[df_h['_s']=='Hoàn thành']))
            r2.metric("🔵 Bình thường",len(df_h[df_h['_s']=='Bình thường']))
            r3.metric("🟡 Nguy cơ",len(df_h[df_h['_s']=='Nguy cơ']))
            r4.metric("🔴 Quá hạn/Trễ",len(df_h[df_h['_s'].isin(['Quá hạn','Trễ tiến độ'])]))
            
            st.dataframe(df_h.drop(columns=['_s']),hide_index=True,width='stretch',
                         column_config={
                             'Tỷ lệ GN': st.column_config.ProgressColumn(
                                 format="%.1f%%", min_value=0, max_value=100
                             )
                         })

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





# ── PAGE: KIỂM DÒ PHÁP LÝ & ĐIỀU KIỆN THANH TOÁN ──
elif page=="⚖️ Kiểm dò pháp lý & Thanh toán":
    st.markdown('<p class="page-title">⚖️ BỘ LỌC KIỂM DÒ PHÁP LÝ & ĐIỀU KIỆN THANH QUYẾT TOÁN SCL</p>', unsafe_allow_html=True)
    st.caption("Căn cứ: QĐ 202/QĐ-HĐTV (EVNHCMC) • TTr 1093/KTAT • Thông báo 2902/TB-PCVT • Nghị định 123/2020/NĐ-CP • Luật Đấu thầu 2023")
    
    row_th, mr, ten, cd = _select_ct("p_legal_sel", header_container)
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


# ── PAGE 2: BẢNG TỔNG HỢP QT ──
elif page=="📋 Bảng tổng hợp QT":
    st.markdown('<p class="page-title">📋 Bảng tổng hợp quyết toán kinh phí SCL</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p2_sel", header_container)
    if row_th is not None:
        ma_ct = str(row_th.get('Mã CT', '')).strip()

        # ── KHUNG TỰ ĐỘNG QUÉT & CẬP NHẬT TỪ HỒ SƠ CÔNG TRÌNH ──
        with st.container(border=True):
            col_info, col_btn = st.columns([3, 1])
            with col_info:
                from dossier_scanner import find_project_folder
                p_folder = find_project_folder(ma_ct, ten)
                if p_folder and os.path.exists(p_folder):
                    rel_p = os.path.relpath(p_folder, BASE_DIR)
                    st.markdown(f"📁 **Thư mục hồ sơ liên kết:** `{rel_p}` ✅")
                    st.caption(f"Công trình: **{ma_ct} - {ten}**. Bấm nút bên phải để quét dữ liệu tự động.")
                else:
                    st.markdown(f"📁 **Thư mục hồ sơ:** Chưa tìm thấy thư mục cho **{ma_ct}** ⚠️")
                    st.caption(f"💡 Anh hãy tạo thư mục: `Ho_so_cong_trinh/{ma_ct}` hoặc `Ho_so_cong_trinh/{ten}` và chép hồ sơ vào đó.")
            with col_btn:
                scan_btn = st.button("🚀 Quét hồ sơ & Cập nhật", type="primary", use_container_width=True, key=f"btn_scan_{ma_ct}")

            if scan_btn:
                with st.spinner("🔄 Đang quét toàn bộ tài liệu hồ sơ và đối soát quy định 202/QĐ-HĐTV..."):
                    from dossier_scanner import scan_and_update_project
                    res_scan = scan_and_update_project(ma_ct, ten)
                    st.session_state[f"scan_res_{ma_ct}"] = res_scan
                    st.rerun()

        # Hiển thị kết quả trích xuất & Kiểm tra pháp lý (nạp tức thì từ cache, không quét lại khi mở trang)
        from dossier_scanner import load_cached_scan_result
        res_scan = st.session_state.get(f"scan_res_{ma_ct}")
        if not res_scan:
            res_scan = load_cached_scan_result(ma_ct)

        if res_scan:
            if res_scan.get('success'):
                up_time = f" *(Cập nhật: {res_scan.get('updated_at', '')})*" if res_scan.get('updated_at') else ""
                st.success(f"✅ {res_scan.get('message')}{up_time}")
                t_scan1, t_scan2 = st.tabs(["📊 Số liệu trích xuất từ hồ sơ", "⚖️ Đánh giá tuân thủ quy định (QĐ 202/QĐ-HĐTV)"])
                with t_scan1:
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        dt_info = res_scan.get('dt_data', {})
                        st.markdown("**1. Số liệu Dự toán:**")
                        if dt_info.get('found'):
                            st.write(f"- Tổng dự toán: **{fmt_full(dt_info.get('tong_dt', 0))}**")
                            st.write(f"- Vật tư thiết bị: {fmt_full(dt_info.get('chi_phi_vttb', 0))}")
                            st.write(f"- Sửa chữa: {fmt_full(dt_info.get('chi_phi_sua_chua', 0))}")
                            st.write(f"- Chi phí khác: {fmt_full(dt_info.get('chi_phi_khac', 0))}")
                            st.write(f"- Dự phòng: {fmt_full(dt_info.get('chi_phi_du_phong', 0))}")
                            st.write(f"- Thu hồi: {fmt_full(dt_info.get('chi_phi_thu_hoi', 0))}")
                            if dt_info.get('so_qd'):
                                st.caption(f"QĐ: {dt_info['so_qd']} ngày {dt_info.get('ngay_qd', '')}")
                        else:
                            st.info("Chưa tìm thấy file Dự toán / QĐ PAKT-DT.")

                    with c2:
                        qt_info = res_scan.get('qt_data', {})
                        st.markdown("**2. Số liệu Quyết toán A-B:**")
                        if qt_info.get('found'):
                            st.write(f"- Tổng quyết toán: **{fmt_full(qt_info.get('tong_qt', 0))}**")
                            st.write(f"- Sửa chữa: {fmt_full(qt_info.get('chi_phi_sua_chua', 0))}")
                            st.write(f"  + Vật liệu: {fmt_full(qt_info.get('vat_lieu', 0))}")
                            st.write(f"  + Nhân công: {fmt_full(qt_info.get('nhan_cong', 0))}")
                            st.write(f"  + Thuế GTGT: {fmt_full(qt_info.get('thue_gtgt', 0))}")
                            st.write(f"- Thu hồi: {fmt_full(qt_info.get('chi_phi_thu_hoi', 0))}")
                        else:
                            st.info("Chưa tìm thấy file Quyết toán A-B trong thư mục.")

                    with c3:
                        ct_info = res_scan.get('contract_data', {})
                        st.markdown("**3. Hợp đồng & Nhà thầu:**")
                        if ct_info.get('found'):
                            st.write(f"- Nhà thầu: **{ct_info.get('nha_thau', '')}**")
                            st.write(f"- Số HĐ: {ct_info.get('so_hd', '')} (ngày {ct_info.get('ngay_hd', '')})")
                            st.write(f"- Giá trị HĐ: {fmt_full(ct_info.get('gia_tri_hd', 0))}")
                            st.write(f"- Hình thức: {ct_info.get('hinh_thuc', 'Thuê ngoài')}")
                            if ct_info.get('ngay_kc') and ct_info.get('ngay_ht'):
                                st.write(f"- Tiến độ: {ct_info.get('ngay_kc')} ➔ {ct_info.get('ngay_ht')}")
                        else:
                            st.info("Chưa có file Hợp đồng.")

                with t_scan2:
                    comp = res_scan.get('compliance', {})
                    if comp:
                        st.markdown(comp.get('summary', ''))
                        for chk in comp.get('checks', []):
                            with st.container(border=True):
                                st.markdown(f"**{chk['icon']} {chk['title']}** — *({chk['rule']})*")
                                st.write(chk['detail'])
                                st.info(f"💡 **Khuyến nghị:** {chk['recommendation']}")
            else:
                st.warning(res_scan.get('message', 'Chưa quét được hồ sơ.'))

        st.divider()

        # Bảng quyết toán kinh phí
        if mr is not None and len(cd)>1:
            sub=cd.iloc[1:][['STT','Tên Công trình','Giá trị Dự toán','Giá trị Q.định phê duyệt QT công trình']].copy()
            sub=sub.rename(columns={'Tên Công trình':'Tên Hạng mục','Giá trị Q.định phê duyệt QT công trình':'Giá trị QT'})
            for c in ['Giá trị Dự toán','Giá trị QT']:
                sub[c]=pd.to_numeric(sub[c],errors='coerce').fillna(0).astype(int)
            sub['Chênh lệch']=sub['Giá trị Dự toán']-sub['Giá trị QT']
            
            # Xác định dòng header (A, B, C, D, E, F, SCL) và dòng tổng
            header_stts = {'A','B','C','D','E','F','SCL'}
            total_stts = {'E','SCL'}
            
            # Build custom HTML table
            html_rows = ''
            for _, r in sub.iterrows():
                stt_val = str(r['STT']).strip()
                ten_hm = str(r['Tên Hạng mục']) if pd.notna(r['Tên Hạng mục']) else ''
                dt_val = fmt_full(r['Giá trị Dự toán'])
                qt_val = fmt_full(r['Giá trị QT'])
                cl_val = fmt_full(r['Chênh lệch'])
                
                row_class = ''
                if stt_val.upper() in total_stts:
                    row_class = ' class="row-total"'
                elif stt_val.upper() in header_stts:
                    row_class = ' class="row-header"'
                
                html_rows += f'''<tr{row_class}>
                    <td class="col-stt">{stt_val}</td>
                    <td class="col-ten">{ten_hm}</td>
                    <td class="col-num">{dt_val}</td>
                    <td class="col-num">{qt_val}</td>
                    <td class="col-num">{cl_val}</td>
                </tr>\n'''
            
            table_html = f'''<div class="qt-table-wrapper">
            <table class="qt-table">
                <thead><tr>
                    <th>STT</th>
                    <th>Tên Hạng mục</th>
                    <th>Giá trị Dự toán</th>
                    <th>Giá trị QT</th>
                    <th>Chênh lệch</th>
                </tr></thead>
                <tbody>{html_rows}</tbody>
            </table></div>'''
            st.markdown(table_html, unsafe_allow_html=True)

            # Xuất Excel
            buf=BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                sub.to_excel(writer, index=False, sheet_name='Quyet_toan')
                workbook = writer.book
                worksheet = writer.sheets['Quyet_toan']
                
                fmt = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'border': 1, 'valign': 'vcenter'})
                fmt_num = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'border': 1, 'valign': 'vcenter', 'num_format': '#,##0'})
                fmt_header = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'bold': True, 'border': 1, 'valign': 'vcenter', 'align': 'center', 'bg_color': '#D9D9D9'})
                
                for col_num, value in enumerate(sub.columns.values):
                    worksheet.write(0, col_num, value, fmt_header)
                
                for i, col in enumerate(sub.columns):
                    if i == 0: width = 6
                    elif i == 1: width = 50
                    else: width = 18
                    worksheet.set_column(i, i, width)
                    
                for row in range(len(sub)):
                    for col in range(len(sub.columns)):
                        val = sub.iloc[row, col]
                        if pd.isna(val): val = ""
                        col_name = sub.columns[col]
                        if col_name in ['Giá trị Dự toán', 'Giá trị QT', 'Chênh lệch']:
                            worksheet.write_number(row + 1, col, val, fmt_num)
                        else:
                            worksheet.write(row + 1, col, val, fmt)

            safe=clean_filename(ten)
            st.markdown(create_download_link(buf.getvalue(), f"Quyet_toan_{safe}.xlsx", "📥 Xuất Excel - Bảng quyết toán", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), unsafe_allow_html=True)
        else:
            st.info("💡 Công trình này chưa có bảng chi tiết quyết toán trong hệ thống. Hãy bấm nút **'🚀 Quét hồ sơ & Cập nhật'** ở trên để hệ thống tự động bóc tách từ hồ sơ và khởi tạo bảng chi tiết!")

# ── PAGE: TỜ TRÌNH DUYỆT QT (MẪU MỚI TỪ THAM KHẢO) ──
elif page=="📝 Tờ trình duyệt QT":
    st.markdown('<p class="page-title">📝 Tờ trình duyệt quyết toán danh mục SCL hoàn thành</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p_ttr_sel", header_container)
    if mr is not None:
        bd=get_cost_breakdown(cd)
        scl_qt=bd.get('SCL',{}).get('qt',0)
        if scl_qt==0 and 'Giá trị Q.định phê duyệt QT công trình' in mr:
            scl_qt=float(mr.get('Giá trị Q.định phê duyệt QT công trình',0))
        thue_qt=bd.get('B.8',{}).get('qt',0)
        if thue_qt==0 and scl_qt>0:
            thue_qt=int(round(scl_qt - scl_qt/1.08))
        truoc_thue=scl_qt - thue_qt if (scl_qt - thue_qt)>0 else scl_qt
        bang_chu_truoc_thue=doc_so_vn(truoc_thue)
        ma=str(mr.get('Mã CT',''))
        so_dt=str(mr.get('Số Dự toán','135/QĐ-PCVT'))
        ngay_dt=_format_date_vn(mr.get('Ngày Dự toán'))
        if '....' in ngay_dt: ngay_dt='16/03/2026'
        now=datetime.datetime.now()

        st.markdown('<p class="section-title">Xem trước nội dung Tờ trình duyệt quyết toán</p>',unsafe_allow_html=True)
        preview_html=f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:15pt;">
                <tr>
                    <td style="width:45%;text-align:center;border:none;vertical-align:top;font-size:11pt;">
                        TỔNG CÔNG TY<br>ĐIỆN LỰC TP HỒ CHÍ MINH<br>
                        <b>CÔNG TY ĐIỆN LỰC VŨNG TÀU</b><br><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /TTr-PCVT<br>
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
            
            <p style="margin-top:20pt;font-weight:bold;">Kính gửi : Tổ thẩm tra phê duyệt quyết toán Sửa chữa lớn</p>
            
            <p style="text-align:justify;line-height:1.6;margin-top:15pt;">
            Căn cứ Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Tổng công ty Điện lực TP.HCM về việc ban hành quy định thực hiện công tác sửa chữa lớn tài sản trong Tổng công ty Điện lực Thành phố Hồ Chí Minh;<br><br>
            Căn cứ quyết định số {so_dt} ngày {ngay_dt} về việc phê duyệt bổ sung danh mục công trình và điều hòa kế hoạch vốn sửa chữa lớn năm {now.year} của Công ty Điện lực Vũng Tàu;<br><br>
            Căn cứ hồ sơ quyết toán công trình {ten}. Đề nghị Tổ thẩm tra phê duyệt quyết toán công trình {ten} mã công trình {ma} với giá trị quyết toán trước thuế: <b>{_fmt_money_dot(truoc_thue)} đồng</b>. (Bằng chữ: {bang_chu_truoc_thue})./.
            </p>
            
            <table style="width:100%;border:none;margin-top:40pt;">
                <tr>
                    <td style="width:50%;border:none;vertical-align:top;font-size:10pt;">
                        <b>Nơi nhận:</b><br>
                        - Như trên;<br>
                        - Lưu: VT, TCKT, HMT.
                    </td>
                    <td style="width:50%;text-align:center;border:none;vertical-align:top;">
                        <b>TM. TỔ THẨM TRA</b><br>
                        <b>TỔ TRƯỞNG</b><br><br><br><br><br>
                        <b>Trần Thanh Hải</b>
                    </td>
                </tr>
            </table>
        </div>
        '''
        data_ttr=export_ttr_duyet_qt_word(mr,cd)
        safe=clean_filename(ten)
        st.markdown(create_download_link(data_ttr, f"TTr_Duyet_QT_{safe}.docx", "📥 Xuất Word - Tờ trình duyệt QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
        st.write('')
        components.html(wrap_preview_html(preview_html), height=700, scrolling=True)

# ── PAGE 3: TMQT ──
elif page=="📄 Thuyết minh QT":
    st.markdown('<p class="page-title">📄 Bảng thuyết minh quyết toán</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p3_sel", header_container)
    if mr is not None:
        bd=get_cost_breakdown(cd)
        scl_dt=bd.get('SCL',{}).get('dt',0);scl_qt=bd.get('SCL',{}).get('qt',0)
        if scl_dt==0 and 'Giá trị Dự toán' in mr: scl_dt=float(mr.get('Giá trị Dự toán',0))
        if scl_qt==0 and 'Giá trị Q.định phê duyệt QT công trình' in mr: scl_qt=float(mr.get('Giá trị Q.định phê duyệt QT công trình',0))
        kh=_safe_int(mr.get('Kế hoạch',0))
        if kh < 100000: kh = kh * 1000000
        if kh == 0: kh = _safe_int(scl_dt)
        dv=str(mr.get('Đơn vị QL','')) if pd.notna(mr.get('Đơn vị QL')) else 'Công ty Cổ phần Sửa chữa ô tô Tiến Phát'
        gc=str(mr.get('Ghi chú','')) if pd.notna(mr.get('Ghi chú')) else 'Thuê ngoài'
        nkc=_format_date_vn(mr.get('Ngày khởi công'))
        nht=_format_date_vn(mr.get('Ngày hoàn thành'))
        if '....' in nkc: nkc = '15/05/2026'
        if '....' in nht: nht = '30/06/2026'
        so_hd=str(mr.get('Số Hợp đồng xây lắp','')) if pd.notna(mr.get('Số Hợp đồng xây lắp')) else ''
        ngay_hd=_format_date_vn(mr.get('Ngày Hợp đồng xây lắp'))
        now=datetime.datetime.now()

        st.markdown('<p class="section-title">Xem trước nội dung Thuyết minh QT</p>',unsafe_allow_html=True)
        klcv=str(mr.get('Khối lượng công việc','')) if pd.notna(mr.get('Khối lượng công việc')) else ''
        noi_dung=str(row_th.get('Nội dung SCL','')) if row_th is not None and pd.notna(row_th.get('Nội dung SCL')) else ''
        if not klcv and noi_dung: klcv=noi_dung

        klcv_html=''
        if klcv:
            for l in [x.strip() for x in klcv.split('\n') if x.strip()]:
                klcv_html += f'<p style="margin-left:20px;line-height:1.4;">{l}</p>\n'
        else:
            klcv_html = '<p style="margin-left:20px;">(Thực hiện đúng theo khối lượng biên bản nghiệm thu hoàn thành công trình)</p>'

        preview_html=f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:10pt;">
                <tr>
                    <td style="width:45%;text-align:center;border:none;vertical-align:top;font-size:11pt;">
                        TỔNG CÔNG TY<br>ĐIỆN LỰC TP HỒ CHÍ MINH<br>
                        <b>CÔNG TY ĐIỆN LỰC VŨNG TÀU</b><br><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /BB-PCVT
                    </td>
                    <td style="width:55%;text-align:center;border:none;vertical-align:top;">
                        <b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b>Độc lập - Tự do - Hạnh Phúc</b><br>
                        -----------------------<br><br>
                        <i>Vũng Tàu, ngày &nbsp;&nbsp;&nbsp; tháng &nbsp;&nbsp;&nbsp; năm {now.year}</i>
                    </td>
                </tr>
            </table>

            <div class="center-title" style="margin-top:15pt;font-weight:bold;font-size:14pt;text-align:center;">
                BẢN THUYẾT MINH QUYẾT TOÁN
            </div>
            
            <p>- Tên gói thầu: Cung cấp vật tư phụ tùng, thi công sửa chữa và mua bảo hiểm</p>
            <p>- Công trình: “<b>{ten}</b>”</p>
            <p>- Giá trị vốn kế hoạch: <b>{_fmt_money_dot(kh)}</b> đồng</p>
            <p>- Hình thức: <b>{gc}</b></p>
            <p>- Tên đơn vị thi công: <b>{dv}</b></p>
            <p>- Giá trị dự toán được duyệt (sau thuế): <b>{_fmt_money_dot(scl_dt)}</b> đồng</p>
            <p>- Thời gian khởi công: <b>{nkc}</b></p>
            <p>- Thời gian hoàn thành: <b>{nht}</b></p>
            <p>- Giá trị quyết toán danh mục hoàn thành: <b>{_fmt_money_dot(scl_qt)}</b> đồng</p>
            
            <p>- Khối lượng công việc chủ yếu đã hoàn thành thay thế sửa chữa, cụ thể:</p>
            {klcv_html}
            
            <p>- Các căn cứ về chế độ để lập quyết toán:</p>
            <p style="margin-left:20px;">+ Căn cứ Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Tổng công ty Điện lực TP.HCM về việc ban hành quy định thực hiện công tác sửa chữa lớn tài sản trong Tổng công ty Điện lực Thành phố Hồ Chí Minh.</p>
            {f'<p style="margin-left:20px;">+ Hợp đồng số: {so_hd} ngày {ngay_hd} giữa Công ty Điện lực Vũng Tàu và {dv}.</p>' if so_hd else ''}
            <p style="margin-left:20px;">+ Bảng kê tổng hợp quyết toán do {dv} lập và được Công ty Điện lực Vũng Tàu thỏa hiệp.</p>
            
            <p>- Phân tích các nhân tố tăng giảm so với dự toán được duyệt: Không có</p>
            <p>- Đánh giá hiệu quả của công việc SCL: Đảm bảo an toàn trong vận hành.</p>
            <p>- Các kiến nghị: Không có ./.</p>
            
            <table style="width:100%;border:none;margin-top:30pt;">
                <tr>
                    <td style="width:50%;border:none;font-size:10pt;vertical-align:top;">
                        <b>Nơi nhận:</b><br>
                        - P/Đ liên quan (để thực hiện);<br>
                        - Lưu: VT, TCKT, HMT.
                    </td>
                    <td style="width:50%;text-align:center;border:none;vertical-align:top;">
                        <b>GIÁM ĐỐC</b><br><br><br><br><br>
                        <b>Nguyễn Ngọc Tuyến</b>
                    </td>
                </tr>
            </table>
        </div>
        '''
        data=export_tmqt_word(mr,cd,noi_dung)
        safe=clean_filename(ten)
        st.markdown(create_download_link(data, f"TMQT_{safe}.docx", "📥 Xuất Word - Thuyết minh QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
        st.write('')
        components.html(wrap_preview_html(preview_html), height=900, scrolling=True)

# ── PAGE 4: PHIẾU THẨM TRA ──
elif page=="🔍 Phiếu thẩm tra":
    st.markdown('<p class="page-title">🔍 Phiếu thẩm tra quyết toán</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p4_sel", header_container)
    if mr is not None:
        gc=str(mr.get('Ghi chú','')) if pd.notna(mr.get('Ghi chú')) else ''
        dv=str(mr.get('Đơn vị QL','')) if pd.notna(mr.get('Đơn vị QL')) else 'Công ty Điện lực Vũng Tàu'
        so_hd=str(mr.get('Số Hợp đồng xây lắp','')) if pd.notna(mr.get('Số Hợp đồng xây lắp')) else ''
        ngay_hd=_format_date_vn(mr.get('Ngày Hợp đồng xây lắp'))
        so_dt=str(mr.get('Số Dự toán','')) if pd.notna(mr.get('Số Dự toán')) else ''
        ngay_dt=_format_date_vn(mr.get('Ngày Dự toán'))
        now=datetime.datetime.now()

        st.markdown('<p class="section-title">Xem trước Phiếu thẩm tra</p>',unsafe_allow_html=True)
        preview_html=f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:10pt;">
                <tr>
                    <td style="width:45%;text-align:center;border:none;vertical-align:top;font-size:11pt;">
                        TỔNG CÔNG TY<br>ĐIỆN LỰC TP HỒ CHÍ MINH<br>
                        <b>CÔNG TY ĐIỆN LỰC VŨNG TÀU</b><br><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /PTT-PCVT
                    </td>
                    <td style="width:55%;text-align:center;border:none;vertical-align:top;">
                        <b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b>Độc lập - Tự do - Hạnh phúc</b><br>
                        -----------------------<br><br>
                        <i>Vũng Tàu, ngày &nbsp;&nbsp;&nbsp; tháng &nbsp;&nbsp;&nbsp; năm {now.year}</i>
                    </td>
                </tr>
            </table>

            <div class="center-title" style="margin-top:15pt;font-weight:bold;font-size:14pt;text-align:center;">
                PHIẾU THẨM TRA QUYẾT TOÁN<br>CÔNG TRÌNH SỬA CHỮA LỚN
            </div>
            
            <p>1. Tên công trình SCL: <b>{ten}</b></p>
            <p>2. Mã công trình: <b>{mr.get('Mã CT','')}</b></p>
            <p>3. Đơn vị quản lý: <b>{dv}</b></p>
            <p>4. Phương thức chọn thầu thực hiện: <b>{gc if gc else 'Thuê ngoài'}</b> {f'(Hợp đồng số {so_hd} ngày {ngay_hd})' if so_hd else ''}</p>
            <p>5. Đơn vị thực hiện: <b>{dv}</b></p>
            <p>6. Căn cứ phê duyệt PAKT-DT: <b>Quyết định số {so_dt} ngày {ngay_dt} của Công ty Điện lực Vũng Tàu</b></p>
            
            <p style="margin-top:15pt;font-weight:bold;">7. Ý KIẾN THẨM TRA CỦA CÁC ĐƠN VỊ THÀNH VIÊN:</p>
            <p>- Phòng Kỹ thuật: Đã kiểm tra khối lượng hoàn công thực tế phù hợp với PAKT-DT được duyệt.</p>
            <p>- Phòng Kế hoạch Vật tư: Đã rà soát chi phí vật tư, hợp đồng và đối chiếu VTTB thu hồi.</p>
            <p>- Phòng Tài chính Kế toán: Đã kiểm tra tính hợp pháp của chứng từ, hóa đơn GTGT và đối chiếu sổ sách kế toán.</p>
            
            <p style="margin-top:15pt;font-weight:bold;">8. KẾT LUẬN VÀ KIẾN NGHỊ:</p>
            <p style="text-align:justify;">Hồ sơ quyết toán công trình SCL đầy đủ tính pháp lý, tuân thủ đúng quy trình quản lý SCL theo Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của EVNHCMC. Kính trình Giám đốc Công ty phê duyệt.</p>
            
            <table style="width:100%;border:none;margin-top:30pt;">
                <tr>
                    <td style="width:33%;text-align:center;border:none;"><b>ĐẠI DIỆN P.KHVT</b></td>
                    <td style="width:33%;text-align:center;border:none;"><b>ĐẠI DIỆN P.TCKT</b></td>
                    <td style="width:33%;text-align:center;border:none;"><b>TỔ TRƯỞNG TỔ THẨM TRA</b></td>
                </tr>
            </table>
        </div>
        '''
        data=export_phieu_tham_tra_word(mr)
        safe=clean_filename(ten)
        st.markdown(create_download_link(data, f"Phieu_tham_tra_{safe}.docx", "📥 Xuất Word - Phiếu thẩm tra QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
        st.write('')
        components.html(wrap_preview_html(preview_html), height=700, scrolling=True)

# ── PAGE 5: BÁO CÁO THẨM TRA (KHỚP Y CHANG THAM KHẢO) ──
elif page=="📜 Báo cáo thẩm tra":
    st.markdown('<p class="page-title">📜 Báo cáo thẩm tra quyết toán công trình SCL</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p5_sel", header_container)
    if mr is not None:
        bd=get_cost_breakdown(cd)
        scl_dt=bd.get('SCL',{}).get('dt',0); scl_qt=bd.get('SCL',{}).get('qt',0)
        if scl_dt==0 and 'Giá trị Dự toán' in mr: scl_dt=float(mr.get('Giá trị Dự toán',0))
        if scl_qt==0 and 'Giá trị Q.định phê duyệt QT công trình' in mr: scl_qt=float(mr.get('Giá trị Q.định phê duyệt QT công trình',0))
        sc_dt=bd.get('B',{}).get('dt',0); sc_qt=bd.get('B',{}).get('qt',0)
        tb_dt=bd.get('A',{}).get('dt',0); tb_qt=bd.get('A',{}).get('qt',0)
        dp_dt=bd.get('D',{}).get('dt',0); dp_qt=bd.get('D',{}).get('qt',0)
        khac_dt=bd.get('C',{}).get('dt',0); khac_qt=bd.get('C',{}).get('qt',0)
        th_qt=bd.get('F',{}).get('qt',0)
        thue_qt=bd.get('B.8',{}).get('qt',0)
        if sc_dt==0 and scl_dt>0: sc_dt=scl_dt
        if sc_qt==0 and scl_qt>0: sc_qt=scl_qt
        if thue_qt==0 and scl_qt>0: thue_qt=int(round(scl_qt - scl_qt/1.08))
        sc_truoc_thue=sc_qt - thue_qt if sc_qt>=thue_qt else sc_qt

        dv5=str(mr.get('Đơn vị QL','')) if pd.notna(mr.get('Đơn vị QL')) else 'Công ty CP sửa chữa Ô tô Tiến Phát'
        gc5=str(mr.get('Ghi chú','')) if pd.notna(mr.get('Ghi chú')) else 'Đấu thầu rộng rãi'
        so_hd5=str(mr.get('Số Hợp đồng xây lắp','')) if pd.notna(mr.get('Số Hợp đồng xây lắp')) else '27-2026/HĐPTV/TP-PCVT'
        ngay_hd5=_format_date_vn(mr.get('Ngày Hợp đồng xây lắp'))
        if '....' in ngay_hd5: ngay_hd5='19/06/2026'
        now=datetime.datetime.now()

        st.markdown('<p class="section-title">Xem trước Báo cáo thẩm tra quyết toán</p>',unsafe_allow_html=True)
        preview5a=f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:10pt;">
                <tr>
                    <td style="width:45%;text-align:center;border:none;vertical-align:top;font-size:11pt;">
                        TỔNG CÔNG TY<br>ĐIỆN LỰC TP HỒ CHÍ MINH<br>
                        <b>CÔNG TY ĐIỆN LỰC VŨNG TÀU</b><br><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /BC-TCKT
                    </td>
                    <td style="width:55%;text-align:center;border:none;vertical-align:top;">
                        <b>CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b>Độc lập - Tự do - Hạnh phúc</b><br>
                        -----------------------<br><br>
                        <i>Vũng Tàu, ngày &nbsp;&nbsp;&nbsp; tháng &nbsp;&nbsp;&nbsp; năm {now.year}</i>
                    </td>
                </tr>
            </table>

            <div class="center-title" style="margin-top:15pt;font-weight:bold;font-size:14pt;text-align:center;">
                BÁO CÁO THẨM TRA QUYẾT TOÁN<br>CÔNG TRÌNH SỬA CHỮA LỚN
            </div>
            
            <p>Căn cứ hồ sơ quyết toán công trình: {ten}.</p>
            <p>Tổ thẩm tra quyết toán công trình sửa chữa lớn Công ty Điện lực Vũng Tàu thẩm tra hồ sơ quyết toán công trình với kết quả như sau:</p>
            
            <p><b><i>I/- Nội dung thẩm tra:</i></b></p>
            <p style="margin-left:20px;">➢ Tên công trình: {ten}.</p>
            <p style="margin-left:20px;">➢ Đơn vị quản lý: Công ty Điện lực Vũng Tàu.</p>
            <p style="margin-left:20px;">➢ Phương thức chọn thầu thực hiện: {gc5}.</p>
            <p style="margin-left:20px;">➢ Hợp đồng số: {so_hd5} ngày {ngay_hd5}.</p>
            <p style="margin-left:20px;">➢ Đơn vị thực hiện: {dv5}.</p>
            
            <p><b>Nội dung cụ thể:</b></p>
            <p style="margin-left:20px;"><b>1/ Phần sửa chữa:</b></p>
            <p style="margin-left:35px;">- Số dự toán &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(sc_dt)} đồng.</p>
            <p style="margin-left:35px;">- Số quyết toán &nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(sc_qt)} đồng.</p>
            <p style="margin-left:35px;">- Số thẩm tra &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(sc_qt)} đồng.</p>
            <p style="margin-left:35px;">- Số chênh lệch &nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(sc_dt - sc_qt)} đồng.</p>
            
            <p style="margin-left:20px;"><b>2/ Phần thiết bị:</b></p>
            <p style="margin-left:35px;">- Số dự toán &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(tb_dt) if tb_dt>0 else ''} đồng.</p>
            <p style="margin-left:35px;">- Số quyết toán &nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(tb_qt) if tb_qt>0 else ''} đồng.</p>
            <p style="margin-left:35px;">- Số thẩm tra &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(tb_qt) if tb_qt>0 else ''} đồng.</p>
            <p style="margin-left:35px;">- Số chênh lệch &nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(tb_dt - tb_qt) if tb_dt>0 else ''} đồng.</p>
            
            <p style="margin-left:20px;"><b>3/ Phần kiến thiết cơ bản khác:</b></p>
            <table style="width:100%;border-collapse:collapse;margin:10pt 0;">
                <tr style="font-weight:bold;text-align:center;background:#f0f0f0;">
                    <td style="border:1px solid #000;padding:5pt;">Nội dung</td>
                    <td style="border:1px solid #000;padding:5pt;width:20%;">Dự toán được duyệt</td>
                    <td style="border:1px solid #000;padding:5pt;width:20%;">Giá trị quyết toán</td>
                    <td style="border:1px solid #000;padding:5pt;width:20%;">Số thẩm tra</td>
                    <td style="border:1px solid #000;padding:5pt;width:20%;">Chênh lệch</td>
                </tr>
                <tr><td style="border:1px solid #000;padding:5pt;">Chi phí thiết kế</td><td style="border:1px solid #000;padding:5pt;"></td><td style="border:1px solid #000;padding:5pt;"></td><td style="border:1px solid #000;padding:5pt;"></td><td style="border:1px solid #000;padding:5pt;"></td></tr>
                <tr><td style="border:1px solid #000;padding:5pt;">Chi phí thẩm định</td><td style="border:1px solid #000;padding:5pt;"></td><td style="border:1px solid #000;padding:5pt;"></td><td style="border:1px solid #000;padding:5pt;"></td><td style="border:1px solid #000;padding:5pt;"></td></tr>
                <tr><td style="border:1px solid #000;padding:5pt;">Chi phí khác</td><td style="border:1px solid #000;padding:5pt;text-align:right;">{_fmt_money_dot(khac_dt) if khac_dt>0 else ''}</td><td style="border:1px solid #000;padding:5pt;text-align:right;">{_fmt_money_dot(khac_qt) if khac_qt>0 else ''}</td><td style="border:1px solid #000;padding:5pt;text-align:right;">{_fmt_money_dot(khac_qt) if khac_qt>0 else ''}</td><td style="border:1px solid #000;padding:5pt;"></td></tr>
                <tr><td style="border:1px solid #000;padding:5pt;">Chi phí dự phòng</td><td style="border:1px solid #000;padding:5pt;text-align:right;">{_fmt_money_dot(dp_dt) if dp_dt>0 else '24.939.900'}</td><td style="border:1px solid #000;padding:5pt;text-align:right;">0</td><td style="border:1px solid #000;padding:5pt;text-align:right;">0</td><td style="border:1px solid #000;padding:5pt;text-align:right;">{_fmt_money_dot(dp_dt) if dp_dt>0 else '0'}</td></tr>
                <tr style="font-weight:bold;"><td style="border:1px solid #000;padding:5pt;">Tổng cộng</td><td style="border:1px solid #000;padding:5pt;text-align:right;">{_fmt_money_dot(dp_dt) if dp_dt>0 else '24.939.900'}</td><td style="border:1px solid #000;padding:5pt;text-align:right;">0</td><td style="border:1px solid #000;padding:5pt;text-align:right;">0</td><td style="border:1px solid #000;padding:5pt;text-align:right;">{_fmt_money_dot(dp_dt) if dp_dt>0 else '0'}</td></tr>
            </table>

            <p><b><i>II/ Kết luận:</i></b></p>
            <p style="margin-left:20px;">Sau khi xem xét thẩm tra hồ sơ, tổ thẩm tra quyết toán chấp thuận tổng giá trị quyết toán công trình nêu trên là: <b>{_fmt_money_dot(scl_qt)}</b> đồng, cụ thể:</p>
            <p style="margin-left:40px;">Chi phí Sửa chữa &nbsp;&nbsp;: &nbsp;&nbsp;<b>{_fmt_money_dot(sc_truoc_thue)}</b> đồng.</p>
            <p style="margin-left:40px;">Chi phí Thiết bị &nbsp;&nbsp;: &nbsp;&nbsp;<b>{_fmt_money_dot(tb_qt)}</b> đồng.</p>
            <p style="margin-left:40px;">Chi phí KTCB khác &nbsp;: &nbsp;&nbsp;<b>{_fmt_money_dot(khac_qt)}</b> đồng.</p>
            <p style="margin-left:40px;">Chi phí VT Thu hồi: &nbsp;&nbsp;<b>{_fmt_money_dot(th_qt)}</b> đồng.</p>
            <p style="margin-left:40px;">Thuế GTGT &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;<b>{_fmt_money_dot(thue_qt)}</b> đồng.</p>
            <p style="margin-left:40px;"><b>Tổng cộng &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(scl_qt)} đồng.</b></p>
            <p style="margin-left:20px;">Kính trình và đề nghị Ông Giám đốc phê duyệt.</p>

            <div style="text-align:center;font-weight:bold;margin:25pt 0 15pt 0;">THÀNH VIÊN TỔ THẨM TRA QUYẾT TOÁN</div>
            <table style="width:100%;border:none;">
                <tr>
                    <td style="width:60%;border:none;line-height:2.0;font-size:11pt;">
                        Hà Thị Mai Hiên (KTT)....................................<br>
                        Nguyễn văn Quyến (TP.QLĐT)............................<br>
                        Nguyễn Mạnh Hiệp (TP.KHVT)............................<br>
                        Đặng Văn Đức (Q.CVP).....................................<br>
                        Đặng Thành Nhân (TTr Công xa).........................<br>
                        Hoàng Minh Tuấn (CV.TCKT)............................
                    </td>
                    <td style="width:40%;text-align:center;border:none;vertical-align:top;">
                        <b>TỔ TRƯỞNG TỔ THẨM TRA</b><br><br><br><br><br><br><br><br>
                        <b>Trần Thanh Hải</b>
                    </td>
                </tr>
            </table>
        </div>
        '''
        data_bc=export_bao_cao_tham_tra_word(mr,cd)
        safe=clean_filename(ten)
        st.markdown(create_download_link(data_bc, f"BC_tham_tra_{safe}.docx", "📥 Xuất Word - Báo cáo thẩm tra", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
        st.write('')
        components.html(wrap_preview_html(preview5a), height=900, scrolling=True)

# ── PAGE 6: QUYẾT ĐỊNH PHÊ DUYỆT (KHỚP Y CHANG THAM KHẢO) ──
elif page=="📜 Quyết định phê duyệt":
    st.markdown('<p class="page-title">📜 Quyết định phê duyệt quyết toán công trình SCL</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p6_sel", header_container)
    if mr is not None:
        bd=get_cost_breakdown(cd)
        scl_qt=bd.get('SCL',{}).get('qt',0)
        if scl_qt==0 and 'Giá trị Q.định phê duyệt QT công trình' in mr:
            scl_qt=float(mr.get('Giá trị Q.định phê duyệt QT công trình',0))
        vttb_qt=bd.get('A',{}).get('qt',0)
        sc_qt=bd.get('B',{}).get('qt',0)
        khac_qt=bd.get('C',{}).get('qt',0)
        th_qt=bd.get('F',{}).get('qt',0)
        thue_qt=bd.get('B.8',{}).get('qt',0)
        if sc_qt==0 and scl_qt>0: sc_qt=scl_qt
        if thue_qt==0 and scl_qt>0: thue_qt=int(round(scl_qt - scl_qt/1.08))
        sc_truoc_thue=sc_qt - thue_qt if sc_qt>=thue_qt else sc_qt

        so_dt=str(mr.get('Số Dự toán','135/QĐ-PCVT'))
        ngay_dt=_format_date_vn(mr.get('Ngày Dự toán'))
        if '....' in ngay_dt: ngay_dt='16/03/2026'
        now=datetime.datetime.now()
        bang_chu=doc_so_vn(scl_qt) if scl_qt>0 else ''

        st.markdown('<p class="section-title">Xem trước Quyết định phê duyệt quyết toán</p>',unsafe_allow_html=True)
        preview5b=f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:10pt;">
                <tr>
                    <td style="width:45%;text-align:center;border:none;vertical-align:top;font-size:11pt;">
                        TỔNG CÔNG TY<br>ĐIỆN LỰC TP HỒ CHÍ MINH<br>
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
            <p>Mã công trình: <b>{mr.get('Mã CT','')}</b></p>
            <p>Chủ đầu tư: <b>Công ty Điện lực Vũng Tàu</b></p>
            <p>Kế hoạch vốn năm: <b>{now.year}</b></p>
            <p>Nguồn vốn: <b>Sửa chữa lớn</b></p>
            
            <div style="text-align:center;font-weight:bold;margin:15pt 0 10pt 0;">GIÁM ĐỐC CÔNG TY ĐIỆN LỰC VŨNG TÀU</div>
            
            <p style="text-align:justify;line-height:1.6;">
            Căn cứ Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Tổng công ty Điện lực TP.HCM về việc ban hành quy định thực hiện công tác sửa chữa lớn tài sản trong Tổng công ty Điện lực Thành phố Hồ Chí Minh;<br>
            Căn cứ quyết định số {so_dt} ngày {ngay_dt} về việc phê duyệt bổ sung danh mục công trình và điều hòa kế hoạch vốn sửa chữa lớn năm {now.year} của Công ty Điện lực Vũng Tàu;<br>
            Căn cứ hồ sơ quyết toán công trình: {ten};<br>
            Căn cứ Báo cáo thẩm tra quyết toán của Tổ thẩm tra về việc phê duyệt quyết toán công trình sửa chữa lớn Công ty Điện lực Vũng Tàu năm {now.year};
            </p>
            
            <div style="text-align:center;font-weight:bold;margin:15pt 0 10pt 0;">QUYẾT ĐỊNH :</div>
            
            <p><b>Điều 1. Phê duyệt quyết toán</b></p>
            <p>Công trình: <b>{ten}</b> với tổng giá trị công trình : <b>{_fmt_money_dot(scl_qt)} đồng</b></p>
            <p>(Bằng chữ: <i>{bang_chu}</i>)</p>
            <p style="margin-left:20px;">Trong đó:</p>
            <p style="margin-left:35px;">Chi phí VTTB &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(vttb_qt) if vttb_qt>0 else ''} đồng</p>
            <p style="margin-left:35px;">Chi phí sửa chữa &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;<b>{_fmt_money_dot(sc_truoc_thue)}</b> đồng</p>
            <p style="margin-left:35px;">Chi phí khác &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(khac_qt) if khac_qt>0 else ''} đồng</p>
            <p style="margin-left:35px;">Vật tư thu hồi &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;{_fmt_money_dot(th_qt) if th_qt>0 else ''} đồng</p>
            <p style="margin-left:35px;">Thuế GTGT &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;&nbsp;<b>{_fmt_money_dot(thue_qt)}</b> đồng</p>
            <p style="margin-left:35px;"><b>Cộng giá trị công trình: &nbsp;&nbsp;{_fmt_money_dot(scl_qt)} đồng</b></p>
            
            <p><b>Điều 2. Nguồn vốn thực hiện công trình</b></p>
            <p>Nguồn vốn: Sửa chữa lớn năm {now.year} của Công ty Điện lực Vũng Tàu.</p>
            
            <p><b>Điều 3. Thực hiện</b></p>
            <p>Phòng Tài chính Kế toán, Phòng Kế hoạch Vật tư, Phòng Quản lý đầu tư, Văn Phòng Công ty Điện lực Vũng Tàu chịu trách nhiệm thi hành quyết định này./.</p>
            
            <table style="width:100%;border:none;margin-top:30pt;">
                <tr>
                    <td style="width:50%;border:none;font-size:10pt;vertical-align:top;">
                        <b>Nơi nhận:</b><br>
                        - P.KHVT, P.QLĐT, VP (để thực hiện);<br>
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
        data_qd=export_qd_phe_duyet_word(mr,cd)
        safe=clean_filename(ten)
        st.markdown(create_download_link(data_qd, f"QD_phe_duyet_{safe}.docx", "📥 Xuất Word - QĐ phê duyệt QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
        st.write('')
        components.html(wrap_preview_html(preview5b), height=900, scrolling=True)

