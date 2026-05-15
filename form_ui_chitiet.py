"""
Module chứa các form chi tiết cho từng hạng mục hồ sơ công trình.
Được gọi từ form_ui.py
"""
import streamlit as st
import pandas as pd
import os
import datetime
from form_module import (
    DANH_MUC_HO_SO, LOAI_HOP_DONG, HINH_THUC_HD,
    HOP_DONG_COLS, PAKT_DT_COLS, KH_DAU_THAU_COLS,
    KQ_DAU_THAU_COLS, VAT_TU_COLS, NGHIEM_THU_QT_COLS,
    load_chitiet_by_ma, save_chitiet_by_ma,
    load_hopdong_list, save_hopdong_list,
    get_category_folder, list_files_in_folder, get_safe_long_path,
    ensure_project_folders, format_num_val, parse_num_val,
    open_file_external
)


def _safe_str(v):
    if pd.isna(v) or v is None: return ''
    return str(v).strip()

def _safe_date(v):
    if pd.isna(v) or v is None: return None
    if isinstance(v, pd.Timestamp): return v.date()
    if isinstance(v, (datetime.date, datetime.datetime)): return v
    return None

def _safe_num(v):
    """Giữ nguyên giá trị số (kể cả thập phân)."""
    if pd.isna(v) or v is None or v == '': return 0
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except: return 0

def _upload_files_to_cat(ten_ct, cat_key, cat_name, wid_key):
    """Upload file vào thư mục hồ sơ tương ứng + hiển thị file có nút Mở/Xóa."""
    cat_folder = get_category_folder(ten_ct, cat_key, cat_name)
    uploaded = st.file_uploader(
        f"📎 Đính kèm file ({cat_key})", 
        accept_multiple_files=True, 
        key=f"up_ct_{cat_key}_{wid_key}",
        label_visibility="visible"
    )
    if uploaded:
        long_folder = get_safe_long_path(cat_folder)
        os.makedirs(long_folder, exist_ok=True)
        for f in uploaded:
            fp = os.path.join(long_folder, f.name)
            with open(fp, 'wb') as out:
                out.write(f.getbuffer())
        st.success(f"✅ Đã lưu {len(uploaded)} file vào mục {cat_key})")
    # Show existing files with Open and Delete buttons
    existing = list_files_in_folder(cat_folder)
    if existing:
        for fn in existing:
            long_folder = get_safe_long_path(cat_folder)
            # Dùng đường dẫn thường cho os.startfile (ứng dụng bên ngoài không hiểu \\?\)
            normal_path = os.path.join(os.path.abspath(cat_folder), fn)
            long_path = os.path.join(long_folder, fn)
            fc1, fc2, fc3 = st.columns([6, 1, 1])
            with fc1:
                st.markdown(f"📎 **{fn}**")
            with fc2:
                if st.button("📂 Mở", key=f"open_{cat_key}_{fn}_{wid_key}"):
                    try:
                        open_file_external(cat_folder, fn)
                    except Exception as e:
                        st.error(f"Không mở được: {e}")
            with fc3:
                if st.button("🗑️ Xóa", key=f"del_ct_{cat_key}_{fn}_{wid_key}"):
                    try:
                        os.remove(long_path)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Không xóa được: {e}")


def render_pakt_dt(ma_ct, ten_ct, wid_key):
    """Form PAKT-DT — liên kết mục c)"""
    with st.expander("📋 PAKT-DT (Phương án khái toán - Dự toán)", expanded=False):
        # Load existing
        df = load_chitiet_by_ma('pakt_dt', ma_ct)
        defaults = {'Số QĐ phê duyệt': '', 'Ngày phê duyệt': None, 'Giá trị dự toán': 0}
        if not df.empty:
            row = df.iloc[0]
            defaults['Số QĐ phê duyệt'] = _safe_str(row.get('Số QĐ phê duyệt'))
            defaults['Ngày phê duyệt'] = _safe_date(row.get('Ngày phê duyệt'))
            defaults['Giá trị dự toán'] = _safe_num(row.get('Giá trị dự toán'))

        c1, c2, c3 = st.columns(3)
        with c1:
            so_qd = st.text_input("Số QĐ phê duyệt PAKT-DT", value=defaults['Số QĐ phê duyệt'], key=f"pakt_soqd_{wid_key}")
        with c2:
            ngay_pd = st.date_input("Ngày phê duyệt", value=defaults['Ngày phê duyệt'], format="DD/MM/YYYY", key=f"pakt_ngay_{wid_key}")
        with c3:
            gt_str = st.text_input("Giá trị dự toán (đồng)", value=format_num_val(defaults['Giá trị dự toán']), key=f"pakt_gt_{wid_key}")
        
        # File upload - mục c
        cat_c = [c for c in DANH_MUC_HO_SO if c[0] == 'c']
        if cat_c:
            _upload_files_to_cat(ten_ct, cat_c[0][0], cat_c[0][1], wid_key)

        return {'Số QĐ phê duyệt': so_qd, 'Ngày phê duyệt': ngay_pd, 'Giá trị dự toán': parse_num_val(gt_str)}


def render_kh_dau_thau(ma_ct, ten_ct, wid_key):
    """Form KH đấu thầu — liên kết mục đ)"""
    with st.expander("📦 Kế hoạch đấu thầu", expanded=False):
        df = load_chitiet_by_ma('kh_dau_thau', ma_ct)
        results = []
        for loai in ['XL', 'TB']:
            st.markdown(f"**Gói thầu {loai}**")
            row_data = df[df['Loại gói'] == loai].iloc[0] if not df.empty and (df['Loại gói'] == loai).any() else pd.Series()
            c1, c2, c3 = st.columns(3)
            with c1:
                soqd = st.text_input(f"Số QĐ KH {loai}", value=_safe_str(row_data.get('Số QĐ phê duyệt KH', '')), key=f"kh_{loai}_soqd_{wid_key}")
            with c2:
                ngay = st.date_input(f"Ngày duyệt {loai}", value=_safe_date(row_data.get('Ngày phê duyệt')), format="DD/MM/YYYY", key=f"kh_{loai}_ngay_{wid_key}")
            with c3:
                gt_str = st.text_input(f"GT gói thầu {loai} (đồng)", value=format_num_val(_safe_num(row_data.get('GT gói thầu', 0))), key=f"kh_{loai}_gt_{wid_key}")
            results.append({'Loại gói': loai, 'Số QĐ phê duyệt KH': soqd, 'Ngày phê duyệt': ngay, 'GT gói thầu': parse_num_val(gt_str)})
        
        # File upload - mục đ
        cat_d = [c for c in DANH_MUC_HO_SO if c[0] == 'đ']
        if cat_d:
            _upload_files_to_cat(ten_ct, cat_d[0][0], cat_d[0][1], wid_key + "_kh")
        
        return results


def render_kq_dau_thau(ma_ct, ten_ct, wid_key):
    """Form Kết quả đấu thầu"""
    with st.expander("🏆 Kết quả đấu thầu", expanded=False):
        df = load_chitiet_by_ma('kq_dau_thau', ma_ct)
        results = []
        for loai in ['XL', 'TB']:
            st.markdown(f"**Kết quả gói {loai}**")
            row_data = df[df['Loại gói'] == loai].iloc[0] if not df.empty and (df['Loại gói'] == loai).any() else pd.Series()
            c1, c2, c3 = st.columns(3)
            with c1:
                soqd = st.text_input(f"Số QĐ, ngày duyệt {loai}", value=_safe_str(row_data.get('Số QĐ phê duyệt KQ', '')), key=f"kq_{loai}_soqd_{wid_key}")
            with c2:
                ngay = st.date_input(f"Ngày duyệt KQ {loai}", value=_safe_date(row_data.get('Ngày phê duyệt')), format="DD/MM/YYYY", key=f"kq_{loai}_ngay_{wid_key}")
            with c3:
                gt_str = st.text_input(f"GT gói thầu {loai} (đồng)", value=format_num_val(_safe_num(row_data.get('GT gói thầu trúng', 0))), key=f"kq_{loai}_gt_{wid_key}")
            c4, c5 = st.columns(2)
            with c4:
                so_hd = st.text_input(f"Số hợp đồng, ngày ký {loai}", value=_safe_str(row_data.get('Số hợp đồng', '')), key=f"kq_{loai}_sohd_{wid_key}")
            with c5:
                gt_hd_str = st.text_input(f"Giá trị hợp đồng {loai} (đồng)", value=format_num_val(_safe_num(row_data.get('Giá trị hợp đồng', 0))), key=f"kq_{loai}_gthd_{wid_key}")
            results.append({
                'Loại gói': loai, 
                'Số QĐ phê duyệt KQ': soqd, 
                'Ngày phê duyệt': ngay, 
                'GT gói thầu trúng': parse_num_val(gt_str),
                'Số hợp đồng': so_hd,
                'Giá trị hợp đồng': parse_num_val(gt_hd_str)
            })
            st.divider()
        
        cat_d = [c for c in DANH_MUC_HO_SO if c[0] == 'đ']
        if cat_d:
            _upload_files_to_cat(ten_ct, cat_d[0][0], cat_d[0][1], wid_key + "_kq")
        
        return results


def render_hop_dong(ma_ct, ten_ct, wid_key):
    """Form Hợp đồng — nhiều HĐ / CT — liên kết mục e)"""
    with st.expander("📝 Hợp đồng", expanded=False):
        # Init session state for contracts
        hd_key = f"hd_list_{ma_ct}"
        if hd_key not in st.session_state:
            df = load_hopdong_list(ma_ct)
            if not df.empty:
                st.session_state[hd_key] = df.to_dict('records')
            else:
                st.session_state[hd_key] = []
        
        hd_list = st.session_state[hd_key]
        
        # Display existing contracts
        if hd_list:
            st.markdown("**Danh sách hợp đồng đã nhập:**")
            for i, hd in enumerate(hd_list):
                with st.container(border=True):
                    col_info, col_del = st.columns([9, 1])
                    with col_info:
                        loai = hd.get('Loại HĐ', '')
                        so_hd = hd.get('Số hợp đồng', '')
                        gt = _safe_num(hd.get('Giá trị HĐ', 0))
                        nha_thau = hd.get('Tên nhà thầu', '')
                        st.markdown(f"**{i+1}. [{loai}]** {so_hd} — GT: {format_num_val(gt)} đ")
                        if nha_thau:
                            st.caption(f"Nhà thầu: {nha_thau}")
                    with col_del:
                        if st.button("🗑️", key=f"del_hd_{i}_{wid_key}"):
                            st.session_state[hd_key].pop(i)
                            st.rerun()
        else:
            st.info("Chưa có hợp đồng nào.")
        
        st.divider()
        st.markdown("**➕ Thêm hợp đồng mới**")
        
        with st.container(border=True):
            st.markdown("**Thông tin hợp đồng**")
            c1, c2 = st.columns(2)
            with c1:
                loai_hd = st.selectbox("Loại HĐ *", LOAI_HOP_DONG, key=f"new_loai_{wid_key}")
                goi_thau = st.text_input("Gói thầu *", key=f"new_goi_{wid_key}")
            with c2:
                hinh_thuc = st.selectbox("Hình thức HĐ *", HINH_THUC_HD, key=f"new_ht_{wid_key}")
                ten_nha_thau = st.text_input("Tên nhà thầu", key=f"new_nthau_{wid_key}")
            
            c3, c4, c5 = st.columns(3)
            with c3:
                so_hd = st.text_input("Số hợp đồng *", key=f"new_sohd_{wid_key}")
            with c4:
                ngay_ky = st.date_input("Ngày ký HĐ *", value=None, format="DD/MM/YYYY", key=f"new_ngayky_{wid_key}")
            with c5:
                ngay_hl = st.date_input("Ngày hiệu lực HĐ", value=None, format="DD/MM/YYYY", key=f"new_ngayhl_{wid_key}")
            
            ten_hd = st.text_input("Tên hợp đồng *", key=f"new_tenhd_{wid_key}")
            
            c6, c7, c8 = st.columns(3)
            with c6:
                gt_hd_str = st.text_input("Giá trị HĐ (đồng) *", value=format_num_val(0), key=f"new_gthd_{wid_key}")
            with c7:
                gt_bl_str = st.text_input("Giá trị bảo lãnh (đồng)", value=format_num_val(0), key=f"new_gtbl_{wid_key}")
            with c8:
                tg_th = st.number_input("Thời gian thực hiện (ngày)", min_value=0, value=0, key=f"new_tgth_{wid_key}")
            
            st.markdown("**Tiến độ thực hiện hợp đồng**")
            c9, c10, c11 = st.columns(3)
            with c9:
                gt_th_str = st.text_input("Giá trị thực hiện HĐ (đồng)", value=format_num_val(0), key=f"new_gtth_{wid_key}")
            with c10:
                so_bb = st.text_input("Số BB nghiệm thu", key=f"new_sobb_{wid_key}")
            with c11:
                ngay_nt = st.date_input("Ngày nghiệm thu", value=None, format="DD/MM/YYYY", key=f"new_ngaynt_{wid_key}")
            
            # File upload - mục e
            cat_e = [c for c in DANH_MUC_HO_SO if c[0] == 'e']
            if cat_e:
                _upload_files_to_cat(ten_ct, cat_e[0][0], cat_e[0][1], wid_key + f"_hd_new")
            
            if st.button("💾 Thêm hợp đồng", type="primary", key=f"btn_add_hd_{wid_key}"):
                if not so_hd.strip():
                    st.error("Vui lòng nhập Số hợp đồng!")
                else:
                    new_hd = {
                        'Mã CT': ma_ct, 'Loại HĐ': loai_hd, 'Gói thầu': goi_thau,
                        'Tên nhà thầu': ten_nha_thau, 'Hình thức HĐ': hinh_thuc,
                        'Số hợp đồng': so_hd, 'Ngày ký HĐ': ngay_ky,
                        'Ngày hiệu lực': ngay_hl, 'Tên hợp đồng': ten_hd,
                        'Giá trị HĐ': parse_num_val(gt_hd_str),
                        'Giá trị bảo lãnh': parse_num_val(gt_bl_str),
                        'Thời gian thực hiện': tg_th,
                        'Giá trị thực hiện HĐ': parse_num_val(gt_th_str),
                        'Số BB nghiệm thu': so_bb, 'Ngày nghiệm thu': ngay_nt,
                    }
                    st.session_state[hd_key].append(new_hd)
                    st.success("✅ Đã thêm hợp đồng!")
                    st.rerun()
        
        return st.session_state[hd_key]


def render_vat_tu(ma_ct, wid_key):
    """Form Vật tư"""
    with st.expander("🔧 Vật tư", expanded=False):
        df = load_chitiet_by_ma('vat_tu', ma_ct)
        defaults = {'TCty cấp': 0, 'ĐV cấp': 0}
        if not df.empty:
            defaults['TCty cấp'] = _safe_num(df.iloc[0].get('TCty cấp', 0))
            defaults['ĐV cấp'] = _safe_num(df.iloc[0].get('ĐV cấp', 0))
        
        c1, c2 = st.columns(2)
        with c1:
            tcty_str = st.text_input("TCty cấp (đồng)", value=format_num_val(defaults['TCty cấp']), key=f"vt_tcty_{wid_key}")
        with c2:
            dv_str = st.text_input("ĐV cấp (đồng)", value=format_num_val(defaults['ĐV cấp']), key=f"vt_dv_{wid_key}")
        
        return {'TCty cấp': parse_num_val(tcty_str), 'ĐV cấp': parse_num_val(dv_str)}


def render_nghiem_thu_qt(ma_ct, ten_ct, wid_key):
    """Form Nghiệm thu & Quyết toán — liên kết mục g, h, i"""
    with st.expander("✅ Nghiệm thu & Quyết toán công trình", expanded=False):
        df = load_chitiet_by_ma('nghiem_thu_qt', ma_ct)
        defaults = {'Ngày nghiệm thu CT': None, 'Giá trị quyết toán CT': 0, 'Ghi chú': ''}
        if not df.empty:
            defaults['Ngày nghiệm thu CT'] = _safe_date(df.iloc[0].get('Ngày nghiệm thu CT'))
            defaults['Giá trị quyết toán CT'] = _safe_num(df.iloc[0].get('Giá trị quyết toán CT', 0))
            defaults['Ghi chú'] = _safe_str(df.iloc[0].get('Ghi chú', ''))
        
        c1, c2, c3 = st.columns(3)
        with c1:
            ngay_nt = st.date_input("Ngày nghiệm thu CT", value=defaults['Ngày nghiệm thu CT'], format="DD/MM/YYYY", key=f"ntqt_ngay_{wid_key}")
        with c2:
            gt_str = st.text_input("Giá trị quyết toán CT (đồng)", value=format_num_val(defaults['Giá trị quyết toán CT']), key=f"ntqt_gt_{wid_key}")
        with c3:
            ghi_chu = st.text_input("Ghi chú", value=defaults['Ghi chú'], key=f"ntqt_gc_{wid_key}")
        
        # File upload - mục g, i
        st.markdown("**Hồ sơ nghiệm thu (mục g):**")
        cat_g = [c for c in DANH_MUC_HO_SO if c[0] == 'g']
        if cat_g:
            _upload_files_to_cat(ten_ct, cat_g[0][0], cat_g[0][1], wid_key + "_g")
        
        st.markdown("**QĐ phê duyệt quyết toán (mục i):**")
        cat_i = [c for c in DANH_MUC_HO_SO if c[0] == 'i']
        if cat_i:
            _upload_files_to_cat(ten_ct, cat_i[0][0], cat_i[0][1], wid_key + "_i")
        
        return {
            'Ngày nghiệm thu CT': ngay_nt,
            'Giá trị quyết toán CT': parse_num_val(gt_str),
            'Ghi chú': ghi_chu
        }


def save_all_chitiet(ma_ct, pakt_data, kh_data, kq_data, hd_list, vt_data, nt_data):
    """Lưu tất cả dữ liệu chi tiết cho 1 CT."""
    # PAKT-DT
    pakt_df = pd.DataFrame([{'Mã CT': ma_ct, **pakt_data}])
    save_chitiet_by_ma('pakt_dt', ma_ct, pakt_df)
    
    # KH đấu thầu
    kh_rows = [{'Mã CT': ma_ct, **r} for r in kh_data]
    save_chitiet_by_ma('kh_dau_thau', ma_ct, pd.DataFrame(kh_rows))
    
    # KQ đấu thầu
    kq_rows = [{'Mã CT': ma_ct, **r} for r in kq_data]
    save_chitiet_by_ma('kq_dau_thau', ma_ct, pd.DataFrame(kq_rows))
    
    # Hợp đồng
    if hd_list:
        hd_df = pd.DataFrame(hd_list)
        hd_df['Mã CT'] = ma_ct
    else:
        hd_df = pd.DataFrame(columns=HOP_DONG_COLS)
    save_hopdong_list(ma_ct, hd_df)
    
    # Vật tư
    vt_df = pd.DataFrame([{'Mã CT': ma_ct, **vt_data}])
    save_chitiet_by_ma('vat_tu', ma_ct, vt_df)
    
    # Nghiệm thu QT
    nt_df = pd.DataFrame([{'Mã CT': ma_ct, **nt_data}])
    save_chitiet_by_ma('nghiem_thu_qt', ma_ct, nt_df)
