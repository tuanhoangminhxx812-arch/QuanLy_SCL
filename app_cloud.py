import streamlit as st
import pandas as pd
import os, re, datetime
from io import BytesIO
from data_helpers import load_tonghop, load_pm092, load_gia_tri_hop_dong
from form_module import load_db_data
from cloud_export import (get_project_section, get_cost_breakdown,
    export_tmqt_word, export_phieu_tham_tra_word,
    export_bao_cao_tham_tra_word, export_qd_phe_duyet_word, _safe_int, _fmt_money_dot)

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
html,body,[class*="st-"]{font-family:'Inter',sans-serif}
.block-container{padding-top:1rem}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#1565C0 0%,#1976D2 40%,#2196F3 100%)}
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3,
[data-testid="stSidebar"] .stMarkdown h5,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stMarkdown span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stRadio label{color:#ffffff!important}
[data-testid="stSidebar"] .stRadio label:hover{background:rgba(255,255,255,.15)!important;border-radius:8px}
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
.a4-preview{background:#ffffff;border:1px solid #e0e0e0;padding:2cm 2.5cm;
max-height:650px;overflow-y:auto;font-family:'Times New Roman',serif;font-size:14pt;
line-height:1.5;color:#000000;box-shadow:0 4px 15px rgba(0,0,0,.15);margin:0 auto;max-width:210mm;}
.a4-preview p{margin:0 0 6pt 0;text-align:justify;}
.a4-preview .center-title{text-align:center;font-weight:bold;font-size:14pt;margin-bottom:15pt;line-height:1.2;}
.a4-preview table{width:100%;border-collapse:collapse;margin:10pt 0;}
.a4-preview td,.a4-preview th{padding:6pt;vertical-align:top;}
div[data-testid="stDataFrame"]{border-radius:10px;overflow:hidden}
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
        v=int(float(v))
        if v>=1e9:return f"{v/1e9:,.2f} tỷ"
        if v>=1e6:return f"{v/1e6:,.1f} tr"
        return f"{v:,}"
    except:return "0"

def fmt_full(v):
    try:return f"{int(float(v)):,}"
    except:return "0"

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
    if status in['Hoàn thành','Nghiệm thu']:return "Hoàn thành","🟢",f"Đã hoàn thành. Giải ngân {r:.1f}%"
    kc_m,kc_y=parse_date_from_text(td,'KC');ht_m,ht_y=parse_date_from_text(td,'HT')
    if ht_m and ht_y:
        ml=(ht_y-cy)*12+(ht_m-cm)
        if ml<0:return "Quá hạn","🔴",f"Trễ {-ml} tháng (HT: {ht_m}/{ht_y}). GN {r:.1f}%"
        if ml<=2 and r<30:return "Nguy cơ","🟡",f"Còn {ml} tháng, GN thấp ({r:.1f}%)"
    if kc_m and kc_y:
        mp=(cy-kc_y)*12+(cm-kc_m)
        if mp>2 and status in['Lập PAKT-Tổng dự toán','Lập kế hoạch đấu thầu']:
            return "Trễ tiến độ","🔴",f"Qua KC {mp} tháng, vẫn '{status}'"
    return "Bình thường","🔵",f"Tiến độ BT. GN {r:.1f}%"

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

# ── Sidebar ──
with st.sidebar:
    st.markdown('<p class="sidebar-logo" style="font-size: 28px !important; line-height: 1.3; margin-bottom: 5px;">Công ty Điện lực Vũng Tàu</p>',unsafe_allow_html=True)
    st.markdown('<p style="color:#ffffff; font-size: 18px; font-weight: 500; margin-top: 0px; margin-bottom: 15px; line-height: 1.4;">Hệ thống quản lý Quyết toán Sửa chữa lớn</p>',unsafe_allow_html=True)
    st.divider()
    page=st.radio("📂 Chuyên mục",
        ["📊 Tổng quan","📋 Thông tin CT","📄 Thuyết minh QT",
         "🔍 Phiếu thẩm tra","📜 BC & QĐ phê duyệt"],
        label_visibility="collapsed")
    st.divider()
    st.caption(f"Cập nhật: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}")



# ── Chọn CT helper ──
def _select_ct(key):
    names=[]
    for _,r in df_th.iterrows():
        ma=str(r.get('Mã CT','')).strip();ten=str(r.get('Tên công trình','')).strip()
        names.append(f"{ma} - {ten}")
    sel=st.selectbox("Chọn công trình:",names,key=key)
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

# ── PAGE 1: TỔNG QUAN ──
if page=="📊 Tổng quan":
    st.markdown('<p class="page-title">📊 Tổng quan các công trình SCL</p>',unsafe_allow_html=True)
    if df_th.empty:
        st.warning("Chưa có dữ liệu Tổng hợp.xlsx")
    else:
        total_ct=len(df_th)
        total_kh=int(df_th['Khái toán'].fillna(0).sum()) if 'Khái toán' in df_th.columns else 0
        total_th=int(df_th['Thực hiện'].fillna(0).sum()) if 'Thực hiện' in df_th.columns else 0
        ty_le=(total_th/total_kh*100) if total_kh>0 else 0

        c1,c2,c3,c4=st.columns(4)
        for col,lbl,val in [(c1,"TỔNG SỐ CÔNG TRÌNH",str(total_ct)),
            (c2,"TỔNG KHÁI TOÁN",fmt_money(total_kh)),
            (c3,"TỔNG THỰC HIỆN",fmt_money(total_th)),
            (c4,"TỶ LỆ GIẢI NGÂN",f"{ty_le:.1f}%")]:
            col.markdown(f'<div class="metric-card"><div class="metric-label">{lbl}</div><div class="metric-val">{val}</div></div>',unsafe_allow_html=True)

        # Health analysis
        now=datetime.datetime.now();cy,cm=now.year,now.month
        hd_list=[]
        for _,r in df_th.iterrows():
            hs,hi,hins=analyze_health(r,cy,cm)
            hd_list.append({'Mã CT':r.get('Mã CT',''),'Tên công trình':r.get('Tên công trình',''),
                'Sức khỏe':f"{hi} {hs}",'Đánh giá':hins,'_s':hs})
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
            st.dataframe(df_h.drop(columns=['_s']),hide_index=True,width='stretch')

        # Charts
        if HAS_PLOTLY:
            st.markdown('<p class="section-title">📈 Biểu đồ trực quan</p>',unsafe_allow_html=True)
            ch1,ch2=st.columns(2)
            with ch1:
                if 'Trạng thái' in df_th.columns:
                    sc=df_th['Trạng thái'].fillna('Chưa XĐ').value_counts()
                    cm2={s:status_color(s) for s in sc.index}
                    fig1=px.pie(values=sc.values,names=sc.index,title="Tỷ trọng trạng thái",
                        color=sc.index,color_discrete_map=cm2,hole=0.4)
                    fig1.update_layout(paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='white',size=12),height=380,margin=dict(t=40,b=20,l=20,r=20))
                    fig1.update_traces(textposition='outside',textinfo='percent+label',textfont_size=11)
                    st.plotly_chart(fig1,width='stretch',key='pie1')
            with ch2:
                if 'Khái toán' in df_th.columns:
                    top=df_th.nlargest(5,'Khái toán').copy()
                    top['KT']=top['Khái toán'].fillna(0)/1e9
                    top['TH']=top['Thực hiện'].fillna(0)/1e9 if 'Thực hiện' in top.columns else 0
                    fig2=go.Figure()
                    fig2.add_trace(go.Bar(name='Khái toán (Tỷ)',x=top['Mã CT'],y=top['KT'],marker_color='#6366f1'))
                    fig2.add_trace(go.Bar(name='Thực hiện (Tỷ)',x=top['Mã CT'],y=top['TH'],marker_color='#f59e0b'))
                    fig2.update_layout(title="Top 5 ngân sách",barmode='group',
                        paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='white',size=12),height=380,margin=dict(t=40,b=20,l=20,r=20))
                    st.plotly_chart(fig2,width='stretch',key='bar1')

        # Table
        st.markdown('<p class="section-title">📋 Bảng số liệu chi tiết</p>',unsafe_allow_html=True)
        dcols=['Mã CT','Tên công trình','Trạng thái','Khái toán','Giá trị HĐ','Thực hiện','Quyết toán']
        dcols=[c for c in dcols if c in df_th.columns]
        dd=df_th[dcols].copy()
        if 'Tên công trình' in dd.columns:
            dd['Tên công trình'] = dd['Tên công trình'].apply(shorten_name)
        for c in ['Khái toán','Giá trị HĐ','Thực hiện','Quyết toán']:
            if c in dd.columns:dd[c]=dd[c].fillna(0).astype(int)
        ccfg={c:st.column_config.NumberColumn(format="%,d") for c in ['Khái toán','Giá trị HĐ','Thực hiện','Quyết toán'] if c in dd.columns}
        st.dataframe(dd,hide_index=True,width='stretch',column_config=ccfg)



# ── PAGE 2: THÔNG TIN CT ──
elif page=="📋 Thông tin CT":
    st.markdown('<p class="page-title">📋 Thông tin chi tiết công trình</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p2_sel")
    if row_th is not None:
        with st.container(border=True):
            c1,c2,c3=st.columns([1.2, 1, 1.8])
            with c1:
                st.write(f"**Tên:** {row_th.get('Tên công trình','')}")
                st.write(f"**Mã CT:** {row_th.get('Mã CT','')}")
                st.write(f"**Trạng thái:** {row_th.get('Trạng thái','')}")
                st.write("")
                if pd.notna(row_th.get('Tiến độ')):
                    st.markdown(f"**Tiến độ:**\n\n{row_th.get('Tiến độ','')}")
            with c2:
                st.write(f"**Khái toán:** {fmt_full(row_th.get('Khái toán',0))} đ")
                st.write(f"**Thực hiện:** {fmt_full(row_th.get('Thực hiện',0))} đ")
                st.write(f"**Quyết toán:** {fmt_full(row_th.get('Quyết toán',0))} đ")
            with c3:
                if pd.notna(row_th.get('Nội dung SCL')):
                    st.markdown(f"**Nội dung SCL:**\n\n{row_th.get('Nội dung SCL','')}")

        # Bảng quyết toán kinh phí
        if mr is not None and len(cd)>1:
            st.markdown('<p class="section-title">Bảng tổng hợp quyết toán kinh phí SCL</p>',unsafe_allow_html=True)
            sub=cd.iloc[1:][['STT','Tên Công trình','Giá trị Dự toán','Giá trị Q.định phê duyệt QT công trình']].copy()
            sub=sub.rename(columns={'Tên Công trình':'Tên Hạng mục','Giá trị Q.định phê duyệt QT công trình':'Giá trị QT'})
            for c in ['Giá trị Dự toán','Giá trị QT']:
                sub[c]=pd.to_numeric(sub[c],errors='coerce').fillna(0).astype(int)
            sub['Chênh lệch']=sub['Giá trị Dự toán']-sub['Giá trị QT']
            st.dataframe(sub,hide_index=True,width='stretch',
                column_config={c:st.column_config.NumberColumn(format="%,d") for c in ['Giá trị Dự toán','Giá trị QT','Chênh lệch']})

            # Xuất Excel
            buf=BytesIO()
            sub.to_excel(buf,index=False,sheet_name='Quyet_toan')
            safe=clean_filename(ten)
            st.markdown(create_download_link(buf.getvalue(), f"Quyet_toan_{safe}.xlsx", "📥 Xuất Excel - Bảng quyết toán", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), unsafe_allow_html=True)

# ── PAGE 3: TMQT ──
elif page=="📄 Thuyết minh QT":
    st.markdown('<p class="page-title">📄 Bảng thuyết minh quyết toán</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p3_sel")
    if mr is not None:
        bd=get_cost_breakdown(cd)
        gt_dt=bd.get('SCL',{}).get('dt',0);gt_qt=bd.get('SCL',{}).get('qt',0)
        kh=_safe_int(mr.get('Kế hoạch',0))
        dv=str(mr.get('Đơn vị QL','')) if pd.notna(mr.get('Đơn vị QL')) else ''
        gc=str(mr.get('Ghi chú','')) if pd.notna(mr.get('Ghi chú')) else ''
        nkc=mr.get('Ngày khởi công');nht=mr.get('Ngày hoàn thành')
        def _fd(d):
            if pd.isna(d) or d is None:return '......'
            if isinstance(d,pd.Timestamp):d=d.date()
            if isinstance(d,(datetime.date,datetime.datetime)):return d.strftime('%d/%m/%Y')
            return str(d)

        st.markdown('<p class="section-title">Xem trước nội dung Thuyết minh QT</p>',unsafe_allow_html=True)
        can_cu=str(mr.get('Căn cứ pháp lý','')) if pd.notna(mr.get('Căn cứ pháp lý')) else ''
        klcv=str(mr.get('Khối lượng công việc','')) if pd.notna(mr.get('Khối lượng công việc')) else ''
        noi_dung=str(row_th.get('Nội dung SCL','')) if row_th is not None and pd.notna(row_th.get('Nội dung SCL')) else ''
        if not klcv and noi_dung: klcv=noi_dung
        chenh=gt_qt-gt_dt
        chenh_txt=f'tăng {_fmt_money_dot(abs(chenh))}' if chenh>0 else (f'giảm {_fmt_money_dot(abs(chenh))}' if chenh<0 else 'bằng dự toán')
        preview_html=f'''
        <div class="a4-preview">
            <div class="center-title">
                BẢNG THUYẾT MINH QUYẾT TOÁN<br>
                CÔNG TRÌNH SỬA CHỮA LỚN HOÀN THÀNH
            </div>
            <p>- Tên danh mục: <b>{ten}</b></p>
            <p>- Mã công trình: <b>{mr.get('Mã CT','')}</b></p>
            <p>- Giá trị vốn kế hoạch: <b>{_fmt_money_dot(kh)}</b> đồng</p>
            <p>- Thuộc kế hoạch vốn sửa chữa lớn năm {datetime.datetime.now().year}</p>
            <p>- Hình thức tự làm hay thuê ngoài: {gc}</p>
            <p>- Tên đơn vị thi công: {dv}</p>
            <p>- Giá trị dự toán được duyệt: <b>{_fmt_money_dot(gt_dt)}</b> đồng</p>
            <p>- Thời gian khởi công: {_fd(nkc)}</p>
            <p>- Thời gian hoàn thành: {_fd(nht)}</p>
            <p>- Giá trị quyết toán danh mục hoàn thành: <b>{_fmt_money_dot(gt_qt)}</b> đồng</p>
            <p>- Khối lượng công việc chủ yếu đã tiến hành: {klcv.replace(chr(10),'<br>')}</p>
            <p>- Các căn cứ về chế độ để lập quyết toán: {can_cu.replace(chr(10),'<br>')}</p>
            <p>- Phân tích các nhân tố tăng giảm: {chenh_txt}</p>
        </div>
        '''
        st.markdown(preview_html,unsafe_allow_html=True)
        st.write('')
        data=export_tmqt_word(mr,cd,noi_dung)
        if data:
            safe=clean_filename(ten)
            st.markdown(create_download_link(data, f"TMQT_{safe}.docx", "📥 Xuất Word - Thuyết minh QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)


# ── PAGE 4: PHIẾU THẨM TRA ──
elif page=="🔍 Phiếu thẩm tra":
    st.markdown('<p class="page-title">🔍 Phiếu thẩm tra quyết toán</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p4_sel")
    if mr is not None:
        gc=str(mr.get('Ghi chú','')) if pd.notna(mr.get('Ghi chú')) else ''
        dv=str(mr.get('Đơn vị QL','')) if pd.notna(mr.get('Đơn vị QL')) else ''
        so_hd=str(mr.get('Số Hợp đồng xây lắp','')) if pd.notna(mr.get('Số Hợp đồng xây lắp')) else ''
        is_tu='tự' in gc.lower() if gc else False

        st.markdown('<p class="section-title">Xem trước Phiếu thẩm tra</p>',unsafe_allow_html=True)
        ngay_hd=mr.get('Ngày Hợp đồng xây lắp')
        ngay_hd_str='......'
        if pd.notna(ngay_hd):
            if isinstance(ngay_hd,pd.Timestamp):ngay_hd=ngay_hd.date()
            if isinstance(ngay_hd,(datetime.date,datetime.datetime)):ngay_hd_str=ngay_hd.strftime('%d/%m/%Y')
        tu_check='✓' if is_tu else '☐'
        thue_check='☐' if is_tu else '✓'
        preview_html=f'''
        <div class="a4-preview">
            <div class="center-title">
                PHIẾU THẨM TRA QUYẾT TOÁN<br>
                CÔNG TRÌNH SỬA CHỮA LỚN
            </div>
            <p>Tên công trình SCL: <b>{ten}</b></p>
            <p>Mã công trình: <b>{mr.get('Mã CT','')}</b></p>
            <p>Đơn vị quản lý: <b>{dv}</b></p>
            <p>Phương thức chọn thầu thực hiện:</p>
            <p style="margin-left:20px;">{tu_check} Tự làm</p>
            <p style="margin-left:20px;">{thue_check} Thuê ngoài (HĐ số {so_hd} ngày {ngay_hd_str})</p>
            <p>Đơn vị thực hiện: <b>{dv}</b></p>
            <br>
            <p style="font-style:italic;color:#666;">(Phần Kết quả kiểm tra — trống, người dùng tự ghi khi in bản cứng)</p>
        </div>
        '''
        st.markdown(preview_html,unsafe_allow_html=True)
        st.write('')
        data=export_phieu_tham_tra_word(mr)
        if data:
            safe=clean_filename(ten)
            st.markdown(create_download_link(data, f"Phieu_tham_tra_{safe}.docx", "📥 Xuất Word - Phiếu thẩm tra QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)

# ── PAGE 5: BÁO CÁO & QĐ ──
elif page=="📜 BC & QĐ phê duyệt":
    st.markdown('<p class="page-title">📜 Báo cáo thẩm tra & Quyết định phê duyệt</p>',unsafe_allow_html=True)
    row_th,mr,ten,cd=_select_ct("p5_sel")
    if mr is not None:
        bd=get_cost_breakdown(cd)
        a_qt=bd.get('A',{}).get('qt',0);b_qt=bd.get('B',{}).get('qt',0)
        c_qt=bd.get('C',{}).get('qt',0);scl_qt=bd.get('SCL',{}).get('qt',0)

        # 5a - Báo cáo thẩm tra
        st.markdown('<p class="section-title">5a. Báo cáo thẩm tra quyết toán</p>',unsafe_allow_html=True)
        dv5=str(mr.get('Đơn vị QL','')) if pd.notna(mr.get('Đơn vị QL')) else ''
        gc5=str(mr.get('Ghi chú','')) if pd.notna(mr.get('Ghi chú')) else ''
        so_hd5=str(mr.get('Số Hợp đồng xây lắp','')) if pd.notna(mr.get('Số Hợp đồng xây lắp')) else ''
        is_tu5='tự' in gc5.lower() if gc5 else False
        a_dt=bd.get('A',{}).get('dt',0);b_dt=bd.get('B',{}).get('dt',0);c_dt=bd.get('C',{}).get('dt',0)
        preview5a=f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:20px;text-align:center;">
                <tr>
                    <td style="width:40%;font-weight:bold;border:none;">TỔNG CÔNG TY ĐIỆN LỰC MIỀN NAM<br>CÔNG TY ĐIỆN LỰC VŨNG TÀU</td>
                    <td style="width:60%;font-weight:bold;border:none;">CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM<br>Độc lập - Tự do - Hạnh phúc</td>
                </tr>
            </table>
            <div class="center-title">
                BÁO CÁO<br>
                Kết quả thẩm tra quyết toán công trình SCL hoàn thành
            </div>
            <p>Tên công trình SCL: <b>{ten}</b></p>
            <p>Đơn vị quản lý: <b>{dv5}</b></p>
            <p>Phương thức chọn thầu thực hiện: <b>{"Tự làm" if is_tu5 else "Thuê ngoài"}</b></p>
            <p>Hợp đồng số: <b>{so_hd5}</b></p>
            <p>Đơn vị thực hiện: <b>{dv5}</b></p>
            <table style="width:100%;border-collapse:collapse;margin:15px 0;">
                <tr style="font-weight:bold;text-align:center;background:#f0f0f0;">
                    <td style="border:1px solid #000;padding:6px;">Hạng mục</td>
                    <td style="border:1px solid #000;padding:6px;">Dự toán</td>
                    <td style="border:1px solid #000;padding:6px;">Quyết toán</td>
                </tr>
                <tr>
                    <td style="border:1px solid #000;padding:6px;">Xây dựng (B)</td>
                    <td style="border:1px solid #000;padding:6px;text-align:right;">{_fmt_money_dot(b_dt)}</td>
                    <td style="border:1px solid #000;padding:6px;text-align:right;">{_fmt_money_dot(b_qt)}</td>
                </tr>
                <tr>
                    <td style="border:1px solid #000;padding:6px;">Thiết bị (A)</td>
                    <td style="border:1px solid #000;padding:6px;text-align:right;">{_fmt_money_dot(a_dt)}</td>
                    <td style="border:1px solid #000;padding:6px;text-align:right;">{_fmt_money_dot(a_qt)}</td>
                </tr>
                <tr>
                    <td style="border:1px solid #000;padding:6px;">KTCB khác (C)</td>
                    <td style="border:1px solid #000;padding:6px;text-align:right;">{_fmt_money_dot(c_dt)}</td>
                    <td style="border:1px solid #000;padding:6px;text-align:right;">{_fmt_money_dot(c_qt)}</td>
                </tr>
                <tr style="font-weight:bold;">
                    <td style="border:1px solid #000;padding:6px;">TỔNG</td>
                    <td style="border:1px solid #000;padding:6px;text-align:right;">{_fmt_money_dot(a_dt+b_dt+c_dt)}</td>
                    <td style="border:1px solid #000;padding:6px;text-align:right;">{_fmt_money_dot(scl_qt)}</td>
                </tr>
            </table>
            <p>Sau khi xem xét thẩm tra hồ sơ, tổ thẩm tra quyết toán chấp thuận tổng giá trị quyết toán công trình nêu trên: <b>{_fmt_money_dot(scl_qt)}</b> đồng</p>
        </div>
        '''
        st.markdown(preview5a,unsafe_allow_html=True)
        st.write('')
        data_bc=export_bao_cao_tham_tra_word(mr,cd)
        if data_bc:
            safe=clean_filename(ten)
            st.markdown(create_download_link(data_bc, f"BC_tham_tra_{safe}.docx", "📥 Xuất Word - Báo cáo thẩm tra", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)

        st.divider()

        # 5b - Quyết định phê duyệt
        st.markdown('<p class="section-title">5b. Quyết định phê duyệt quyết toán</p>',unsafe_allow_html=True)
        from form_module import doc_so_vn
        bang_chu=doc_so_vn(scl_qt) if scl_qt>0 else ''
        preview5b=f'''
        <div class="a4-preview">
            <table style="width:100%;border:none;margin-bottom:20px;text-align:center;">
                <tr>
                    <td style="width:40%;font-weight:bold;border:none;">TỔNG CÔNG TY ĐIỆN LỰC MIỀN NAM<br>CÔNG TY ĐIỆN LỰC VŨNG TÀU</td>
                    <td style="width:60%;font-weight:bold;border:none;">CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM<br>Độc lập - Tự do - Hạnh phúc</td>
                </tr>
            </table>
            <div class="center-title">
                QUYẾT ĐỊNH<br>
                V/v Phê duyệt quyết toán công trình sửa chữa lớn
            </div>
            <p style="text-align:center;font-weight:bold;margin-bottom:15pt;">GIÁM ĐỐC CÔNG TY ĐIỆN LỰC VŨNG TÀU</p>
            <p><b>Điều 1:</b> Phê duyệt quyết toán công trình: <b>{ten}</b></p>
            <p>với tổng giá trị: <b>{_fmt_money_dot(scl_qt)}</b> đồng</p>
            <p>(Bằng chữ: <i>{bang_chu}</i>)</p>
            <p style="margin-left:20px;">- Chi phí thiết bị: {_fmt_money_dot(a_qt)} đồng</p>
            <p style="margin-left:20px;">- Chi phí xây dựng: {_fmt_money_dot(b_qt)} đồng</p>
            <p style="margin-left:20px;">- KTCB khác: {_fmt_money_dot(c_qt)} đồng</p>
            <p><b>Điều 2:</b> Nguồn vốn thực hiện công trình: Sửa chữa lớn của Điện lực Vũng Tàu.</p>
        </div>
        '''
        st.markdown(preview5b,unsafe_allow_html=True)
        st.write('')
        data_qd=export_qd_phe_duyet_word(mr,cd)
        if data_qd:
            safe=clean_filename(ten)
            st.markdown(create_download_link(data_qd, f"QD_phe_duyet_{safe}.docx", "📥 Xuất Word - QĐ phê duyệt QT", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), unsafe_allow_html=True)
