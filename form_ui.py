import streamlit as st
import pandas as pd
import os
from form_module import (
    ALL_COLUMNS, DB_FILE, DANH_MUC_HO_SO, 
    load_db_data, format_num_val, parse_num_val, doc_so_vn,
    ensure_project_folders, get_category_folder, list_files_in_folder,
    get_safe_long_path, open_file_external
)
from form_ui_chitiet import (
    render_pakt_dt, render_kh_dau_thau, render_kq_dau_thau,
    render_hop_dong, render_vat_tu, render_nghiem_thu_qt,
    save_all_chitiet
)

def render_full_form(selected_edit_ct, sel_ma, new_tt, new_tiendo, new_noidung):
    db_df_tab1 = load_db_data()
    
    defaults = {
        'STT': "I", 'Tên Công trình': selected_edit_ct, 'Mã CT': sel_ma, 'Kế hoạch': 0, 'Số Phương án': "", 'Ngày Phương án': None,
        'Giá trị Phương án': 0, 'Số Dự toán': "", 'Ngày Dự toán': None, 'Giá trị Dự toán': 0, 'Số Hợp đồng thiết kế': "",
        'Ngày Hợp đồng thiết kế': None, 'Giá trị Hợp đồng thiết kế': 0, 'Số Hợp đồng giám sát': "", 'Ngày Hợp đồng giám sát': None,
        'Giá trị Hợp đồng giám sát': 0, 'Số Hợp đồng xây lắp': "", 'Ngày Hợp đồng xây lắp': None, 'Giá trị Hợp đồng xây lắp': 0,
        'Giá trị phát sinh': 0, 'Giá trị VT thừa': 0, 'Giá trị VTTH': 0, 'Số Q.định phê duyệt QT công trình': "",
        'Ngày Q.định phê duyệt QT công trình': None, 'Giá trị Q.định phê duyệt QT công trình': 0, 'Ghi chú': "", 'Đơn vị QL': "",
        'Căn cứ pháp lý': "", 'Khối lượng công việc': "", 'Ngày khởi công': None, 'Ngày hoàn thành': None
    }

    sub_items_columns = ['STT', 'Tên Hạng mục', 'Giá trị Dự toán', 'Giá trị quyết toán', 'Chênh lệch']
    
    # Load data from DB if exists
    # Old app used 'Tên Công trình' to find the record, fallback to Mã CT just in case
    start_indices = db_df_tab1.index[db_df_tab1['Tên Công trình'] == selected_edit_ct].tolist()
    if not start_indices:
        start_indices = db_df_tab1.index[db_df_tab1['Mã CT'].astype(str).str.strip() == sel_ma].tolist()
    if start_indices:
        start_idx = start_indices[0]
        end_idx = len(db_df_tab1)
        for i in range(start_idx + 1, len(db_df_tab1)):
            val = str(db_df_tab1.at[i, 'STT']).strip().upper()
            if val in ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']:
                end_idx = i
                break
        
        active_df = db_df_tab1.iloc[start_idx:end_idx]
        main_row_data = active_df.iloc[0].to_dict()
        for k in defaults.keys():
            if pd.notna(main_row_data.get(k)):
                defaults[k] = main_row_data[k]
                if isinstance(defaults[k], pd.Timestamp):
                    defaults[k] = defaults[k].date()
        
        if st.session_state.get('last_edit_ct') != selected_edit_ct:
            sub_rows = active_df.iloc[1:].copy()
            sub_rows = sub_rows[['STT', 'Tên Công trình', 'Giá trị Dự toán', 'Giá trị Q.định phê duyệt QT công trình', 'Ghi chú']]
            sub_rows = sub_rows.rename(columns={'Tên Công trình': 'Tên Hạng mục', 'Giá trị Q.định phê duyệt QT công trình': 'Giá trị quyết toán', 'Ghi chú': 'Chênh lệch'})
            sub_rows['Giá trị Dự toán'] = sub_rows['Giá trị Dự toán'].fillna(0).astype(int)
            sub_rows['Giá trị quyết toán'] = sub_rows['Giá trị quyết toán'].fillna(0).astype(int)
            for col in ['STT', 'Tên Hạng mục', 'Chênh lệch']:
                sub_rows[col] = sub_rows[col].fillna("")
                
            st.session_state.sub_df = sub_rows
            st.session_state.last_edit_ct = selected_edit_ct
    else:
        if st.session_state.get('last_edit_ct') != selected_edit_ct:
            if 'sub_df' in st.session_state:
                del st.session_state['sub_df']
            st.session_state.last_edit_ct = selected_edit_ct

    wid_key = sel_ma
    
    with st.expander("Thông tin cơ bản", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            stt = st.text_input("STT (VD: I, II)", value=str(defaults['STT']), key=f"stt_{wid_key}")
            ten_ct = st.text_input("Tên Công trình (* Bắt buộc)", value=str(defaults['Tên Công trình']), key=f"ten_ct_{wid_key}")
            don_vi_ql = st.text_input("Đơn vị Quản lý", value=str(defaults['Đơn vị QL']), key=f"dvql_{wid_key}")
        with col2:
            ma_ct = st.text_input("Mã CT", value=str(defaults['Mã CT']), key=f"ma_ct_{wid_key}")
            ke_hoach_str = st.text_input("Kế hoạch (Giá trị)", value=format_num_val(defaults['Kế hoạch']), key=f"kehoach_{wid_key}")
            ke_hoach = parse_num_val(ke_hoach_str)
            so_tien_chu = st.text_input("Số tiền bằng chữ", value=doc_so_vn(ke_hoach), key=f"sotien_{wid_key}")
        with col3:
            ghi_chu = st.text_input("Ghi chú", value=str(defaults['Ghi chú']), key=f"ghichu_{wid_key}")
            col_ngay1, col_ngay2 = st.columns(2)
            with col_ngay1:
                ngay_khoi_cong = st.date_input("Ngày khởi công", value=defaults['Ngày khởi công'], format="DD/MM/YYYY", key=f"ngaykc_{wid_key}")
            with col_ngay2:
                ngay_hoan_thanh = st.date_input("Ngày hoàn thành", value=defaults['Ngày hoàn thành'], format="DD/MM/YYYY", key=f"ngayht_{wid_key}")
        
        can_cu_phap_ly = st.text_area("Căn cứ pháp lý", value=str(defaults['Căn cứ pháp lý']), height=100, key=f"ccpl_{wid_key}")
        khoi_luong_cv = st.text_area("Khối lượng công việc", value=str(defaults['Khối lượng công việc']), height=100, key=f"klcv_{wid_key}")
    
    with st.expander("Thông tin Phương án, Dự toán và Phê duyệt QT"):
        col4, col5, col6 = st.columns(3)
        with col4:
            st.markdown("**Phương án**")
            so_pa = st.text_input("Số Phương án", value=str(defaults['Số Phương án']), key=f"sopa_{wid_key}")
            ngay_pa = st.date_input("Ngày Phương án", value=defaults['Ngày Phương án'], format="DD/MM/YYYY", key=f"ngaypa_{wid_key}")
            gt_pa_str = st.text_input("Giá trị Phương án", value=format_num_val(defaults['Giá trị Phương án']), key=f"gtpa_{wid_key}")
            gt_pa = parse_num_val(gt_pa_str)
        with col5:
            st.markdown("**Dự toán**")
            so_dt = st.text_input("Số Dự toán", value=str(defaults['Số Dự toán']), key=f"sodt_{wid_key}")
            ngay_dt = st.date_input("Ngày Dự toán", value=defaults['Ngày Dự toán'], format="DD/MM/YYYY", key=f"ngaydt_{wid_key}")
            gt_dt_str = st.text_input("Giá trị Dự toán", value=format_num_val(defaults['Giá trị Dự toán']), key=f"gtdt_{wid_key}")
            gt_dt = parse_num_val(gt_dt_str)
        with col6:
            st.markdown("**QĐ Phê duyệt QT CT**")
            so_qd = st.text_input("Số Q.định phê duyệt QT", value=str(defaults['Số Q.định phê duyệt QT công trình']), key=f"soqd_{wid_key}")
            ngay_qd = st.date_input("Ngày Q.định phê duyệt QT", value=defaults['Ngày Q.định phê duyệt QT công trình'], format="DD/MM/YYYY", key=f"ngayqd_{wid_key}")
            gt_qd_str = st.text_input("Giá trị Q.định phê duyệt QT", value=format_num_val(defaults['Giá trị Q.định phê duyệt QT công trình']), key=f"gtqd_{wid_key}")
            gt_qd = parse_num_val(gt_qd_str)
            
    with st.expander("Thông tin Hợp đồng & Vật tư khác (Tùy chọn)"):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**HĐ Thiết kế**")
            so_hdtk = st.text_input("Số HĐ TK", value=str(defaults['Số Hợp đồng thiết kế']), key=f"sohdtk_{wid_key}")
            ngay_hdtk = st.date_input("Ngày HĐ TK", value=defaults['Ngày Hợp đồng thiết kế'], format="DD/MM/YYYY", key=f"ngayhdtk_{wid_key}")
            gt_hdtk_str = st.text_input("Giá trị HĐ TK", value=format_num_val(defaults['Giá trị Hợp đồng thiết kế']), key=f"gthdtk_{wid_key}")
            gt_hdtk = parse_num_val(gt_hdtk_str)
        with c2:
            st.markdown("**HĐ Giám sát**")
            so_hdgs = st.text_input("Số HĐ GS", value=str(defaults['Số Hợp đồng giám sát']), key=f"sohdgs_{wid_key}")
            ngay_hdgs = st.date_input("Ngày HĐ GS", value=defaults['Ngày Hợp đồng giám sát'], format="DD/MM/YYYY", key=f"ngayhdgs_{wid_key}")
            gt_hdgs_str = st.text_input("Giá trị HĐ GS", value=format_num_val(defaults['Giá trị Hợp đồng giám sát']), key=f"gthdgs_{wid_key}")
            gt_hdgs = parse_num_val(gt_hdgs_str)
        with c3:
            st.markdown("**HĐ Xây lắp**")
            so_hdxl = st.text_input("Số HĐ XL", value=str(defaults['Số Hợp đồng xây lắp']), key=f"sohdxl_{wid_key}")
            ngay_hdxl = st.date_input("Ngày HĐ XL", value=defaults['Ngày Hợp đồng xây lắp'], format="DD/MM/YYYY", key=f"ngayhdxl_{wid_key}")
            gt_hdxl_str = st.text_input("Giá trị HĐ XL", value=format_num_val(defaults['Giá trị Hợp đồng xây lắp']), key=f"gthdxl_{wid_key}")
            gt_hdxl = parse_num_val(gt_hdxl_str)
        
        c4, c5, c6 = st.columns(3)
        with c4:
            gt_ps_str = st.text_input("Giá trị phát sinh", value=format_num_val(defaults['Giá trị phát sinh']), key=f"gtps_{wid_key}")
            gt_ps = parse_num_val(gt_ps_str)
        with c5:
            gt_vtt_str = st.text_input("Giá trị VT thừa", value=format_num_val(defaults['Giá trị VT thừa']), key=f"gtvtt_{wid_key}")
            gt_vtt = parse_num_val(gt_vtt_str)
        with c6:
            gt_vtth_str = st.text_input("Giá trị VTTH", value=format_num_val(defaults['Giá trị VTTH']), key=f"gtvtth_{wid_key}")
            gt_vtth = parse_num_val(gt_vtth_str)

    # ============================================================
    # PHẦN CHI TIẾT CÔNG TRÌNH MỚI
    # ============================================================
    st.markdown("##### 📋 Chi tiết hồ sơ theo hạng mục")
    
    pakt_data = render_pakt_dt(sel_ma, ten_ct, wid_key)
    kh_data = render_kh_dau_thau(sel_ma, ten_ct, wid_key)
    kq_data = render_kq_dau_thau(sel_ma, ten_ct, wid_key)
    hd_data = render_hop_dong(sel_ma, ten_ct, wid_key)
    vt_data = render_vat_tu(sel_ma, wid_key)
    nt_data = render_nghiem_thu_qt(sel_ma, ten_ct, wid_key)

    # ============================================================
    # PHẦN HỒ SƠ CÔNG TRÌNH (giữ nguyên)
    # ============================================================
    with st.expander("📂 Hồ sơ công trình", expanded=False):
        if not ten_ct.strip():
            st.warning("⚠️ Vui lòng nhập **Tên Công trình** ở phần Thông tin cơ bản trước khi quản lý hồ sơ.")
        else:
            project_dir = ensure_project_folders(ten_ct)
            st.caption(f"📁 Thư mục lưu trữ: `{project_dir}`")
            st.divider()
            
            for cat_key, cat_name in DANH_MUC_HO_SO:
                cat_folder = get_category_folder(ten_ct, cat_key, cat_name)
                existing_files = list_files_in_folder(cat_folder)
                has_files = len(existing_files) > 0
                
                col_name, col_upload = st.columns([7, 3])
                with col_name:
                    status_icon = "✅" if has_files else "⬜"
                    st.markdown(f"**{status_icon} {cat_key}) {cat_name}**")
                
                with col_upload:
                    uploaded_files = st.file_uploader("Thêm file", accept_multiple_files=True, key=f"up_{cat_key}_{wid_key}", label_visibility="collapsed")
                    if uploaded_files:
                        long_cat_folder = get_safe_long_path(cat_folder)
                        os.makedirs(long_cat_folder, exist_ok=True)
                        for uploaded_file in uploaded_files:
                            file_path = os.path.join(long_cat_folder, uploaded_file.name)
                            with open(file_path, 'wb') as f:
                                f.write(uploaded_file.getbuffer())
                        st.success(f"Đã lưu {len(uploaded_files)} file")
                        st.rerun()
                
                if existing_files:
                    with st.container():
                        for file_name in existing_files:
                            long_cat_folder = get_safe_long_path(cat_folder)
                            long_path = os.path.join(long_cat_folder, file_name)
                            normal_path = os.path.join(os.path.abspath(cat_folder), file_name)
                            file_size = os.path.getsize(long_path)
                            size_str = f"{file_size/1024:.1f} KB" if file_size < 1024*1024 else f"{file_size/1024/1024:.1f} MB"
                            fc1, fc2, fc3, fc4 = st.columns([5, 2, 1, 1])
                            with fc1: st.markdown(f"📎 **{file_name}**")
                            with fc2: st.caption(f"({size_str})")
                            with fc3:
                                if st.button("📂 Mở", key=f"hs_open_{cat_key}_{file_name}_{wid_key}"):
                                    try:
                                        open_file_external(cat_folder, file_name)
                                    except Exception as e:
                                        st.error(f"Lỗi: {e}")
                            with fc4:
                                if st.button("🗑️", key=f"hs_del_{cat_key}_{file_name}_{wid_key}"):
                                    os.remove(long_path)
                                    st.rerun()
                st.divider()

    st.markdown("##### BẢNG TỔNG HỢP QUYẾT TOÁN KINH PHÍ SỬA CHỮA LỚN")
    
    if 'sub_df' not in st.session_state:
        initial_data = [
            {"STT": "A", "Tên Hạng mục": "CHI PHÍ VẬT TƯ, THIẾT BỊ (sau thuế)", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "A.1", "Tên Hạng mục": "Chi phí thiết bị", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "A.1.1", "Tên Hạng mục": "Thiết bị nhập khẩu", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "A.1.2", "Tên Hạng mục": "VT A cấp", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "A.1.3", "Tên Hạng mục": "Chi phí tháo dỡ, lắp đặt", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "A.1.4", "Tên Hạng mục": "Chi phí thí nghiệm, hiệu chỉnh", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "A.2", "Tên Hạng mục": "Chi phí vật tư", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "A.3", "Tên Hạng mục": "Thuế GTGT", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B", "Tên Hạng mục": "CHI PHÍ SỬA CHỮA", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.1", "Tên Hạng mục": "Chi phí vật liệu", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.1.1", "Tên Hạng mục": "Vật liệu phần không áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.1.2", "Tên Hạng mục": "Vật liệu phần áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.1.3", "Tên Hạng mục": "Chênh lệch giá vật liệu phần áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.1.4", "Tên Hạng mục": "Vật liệu phụ trong SCL thiết bị", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.2", "Tên Hạng mục": "Chi phí nhân công", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.2.1", "Tên Hạng mục": "Chi phí nhân công phần không áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.2.2", "Tên Hạng mục": "Chi phí nhân công phần áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.3", "Tên Hạng mục": "Chi phí máy thi công", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.3.1", "Tên Hạng mục": "Chi phí máy thi công phần không áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.3.2", "Tên Hạng mục": "Chi phí máy thi công phần áp dụng đơn giá XDCB", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.4", "Tên Hạng mục": "Chi phí làm đêm, làm thêm giờ", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.5", "Tên Hạng mục": "Chi phí chung", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.6", "Tên Hạng mục": "Thu nhập chịu thuế tính trước", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.7", "Tên Hạng mục": "Giá trị sửa chữa trước thuế", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "B.8", "Tên Hạng mục": "Thuế GTGT", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "C", "Tên Hạng mục": "CHI PHÍ KHÁC (sau thuế)", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "C.1", "Tên Hạng mục": "Chi phí giám sát thi công xây dựng", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "C.2", "Tên Hạng mục": "Chi phí giám sát lắp đặt thiết bị", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "C.3", "Tên Hạng mục": "Chi phí bảo hiểm công trình", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "C.4", "Tên Hạng mục": "Chi phí thẩm tra - phê duyệt quyết toán", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "C.5", "Tên Hạng mục": "Vận chuyển VTTB A cấp đến công trường", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "C.6", "Tên Hạng mục": "Thuế GTGT", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "D", "Tên Hạng mục": "CHI PHÍ DỰ PHÒNG", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "E", "Tên Hạng mục": "Tổng giá trị sau thuế", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "E.1", "Tên Hạng mục": "Tổng giá trị trước thuế", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "E.2", "Tên Hạng mục": "Thuế GTGT", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "F", "Tên Hạng mục": "GIÁ TRỊ VẬT TƯ THU HỒI", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0},
            {"STT": "SCL", "Tên Hạng mục": "CHI PHÍ SCL", "Giá trị Dự toán": 0, "Giá trị quyết toán": 0, "Chênh lệch": 0}
        ]
        st.session_state.sub_df = pd.DataFrame(initial_data, columns=sub_items_columns)
    
    col_config = {
        "Giá trị Dự toán": st.column_config.NumberColumn(format="%,d", step=1),
        "Giá trị quyết toán": st.column_config.NumberColumn(format="%,d", step=1),
        "Chênh lệch": st.column_config.NumberColumn(format="%,d", disabled=True)
    }
    
    st.session_state.sub_df['Chênh lệch'] = pd.to_numeric(st.session_state.sub_df['Chênh lệch'], errors='coerce').fillna(0).astype(int)
    edited_sub_df = st.data_editor(st.session_state.sub_df, num_rows="dynamic", width="stretch", column_config=col_config, key=f"editor_{wid_key}")

    # TÍNH TOÁN
    calculated_df = edited_sub_df.copy()
    for col in ['Giá trị Dự toán', 'Giá trị quyết toán']:
        calculated_df[col] = pd.to_numeric(calculated_df[col], errors='coerce').fillna(0).astype(int)
    calculated_df['Chênh lệch'] = calculated_df['Giá trị Dự toán'] - calculated_df['Giá trị quyết toán']

    changed = [False]
    def set_val(target_stt, val_dt, val_qt):
        mask = calculated_df['STT'].astype(str).str.strip() == target_stt
        if mask.any():
            idx = mask.idxmax()
            if int(calculated_df.at[idx, 'Giá trị Dự toán']) != int(val_dt):
                calculated_df.at[idx, 'Giá trị Dự toán'] = int(val_dt)
                changed[0] = True
            if int(calculated_df.at[idx, 'Giá trị quyết toán']) != int(val_qt):
                calculated_df.at[idx, 'Giá trị quyết toán'] = int(val_qt)
                changed[0] = True

    def get_val(target_stt):
        mask = calculated_df['STT'].astype(str).str.strip() == target_stt
        if mask.any():
            idx = mask.idxmax()
            return int(calculated_df.at[idx, 'Giá trị Dự toán']), int(calculated_df.at[idx, 'Giá trị quyết toán'])
        return 0, 0

    def get_sum(stt_list):
        s_dt, s_qt = 0, 0
        for stt in stt_list:
            dt, qt = get_val(stt)
            s_dt += dt
            s_qt += qt
        return s_dt, s_qt

    val_A1_dt, val_A1_qt = get_sum(['A.1.1', 'A.1.2', 'A.1.3', 'A.1.4'])
    set_val('A.1', val_A1_dt, val_A1_qt)
    val_A2_dt, val_A2_qt = get_val('A.2')
    val_A3_dt = int(round((val_A1_dt + val_A2_dt) * 0.10))
    val_A3_qt = int(round((val_A1_qt + val_A2_qt) * 0.10))
    set_val('A.3', val_A3_dt, val_A3_qt)
    val_A_dt = val_A1_dt + val_A2_dt + val_A3_dt
    val_A_qt = val_A1_qt + val_A2_qt + val_A3_qt
    set_val('A', val_A_dt, val_A_qt)

    val_B1_dt, val_B1_qt = get_sum(['B.1.1', 'B.1.2', 'B.1.3', 'B.1.4'])
    set_val('B.1', val_B1_dt, val_B1_qt)
    val_B2_dt, val_B2_qt = get_sum(['B.2.1', 'B.2.2'])
    set_val('B.2', val_B2_dt, val_B2_qt)
    val_B3_dt, val_B3_qt = get_sum(['B.3.1', 'B.3.2'])
    set_val('B.3', val_B3_dt, val_B3_qt)
    val_B4_dt, val_B4_qt = get_val('B.4')
    val_B5_dt, val_B5_qt = get_val('B.5')
    val_B6_dt, val_B6_qt = get_val('B.6')

    val_B7_dt = sum([val_B1_dt, val_B2_dt, val_B3_dt, val_B4_dt, val_B5_dt, val_B6_dt])
    val_B7_qt = sum([val_B1_qt, val_B2_qt, val_B3_qt, val_B4_qt, val_B5_qt, val_B6_qt])
    set_val('B.7', val_B7_dt, val_B7_qt)
    val_B8_dt = int(round(val_B7_dt * 0.08))
    val_B8_qt = int(round(val_B7_qt * 0.08))
    set_val('B.8', val_B8_dt, val_B8_qt)
    val_B_dt = val_B7_dt + val_B8_dt
    val_B_qt = val_B7_qt + val_B8_qt
    set_val('B', val_B_dt, val_B_qt)

    val_C1_dt, val_C1_qt = get_val('C.1')
    val_C2_dt, val_C2_qt = get_val('C.2')
    val_C3_dt, val_C3_qt = get_val('C.3')
    val_C4_dt, val_C4_qt = get_val('C.4')
    val_C5_dt, val_C5_qt = get_val('C.5')
    sum_C_1_5_dt = val_C1_dt + val_C2_dt + val_C3_dt + val_C4_dt + val_C5_dt
    sum_C_1_5_qt = val_C1_qt + val_C2_qt + val_C3_qt + val_C4_qt + val_C5_qt
    val_C6_dt = int(round(sum_C_1_5_dt * 0.08))
    val_C6_qt = int(round(sum_C_1_5_qt * 0.08))
    set_val('C.6', val_C6_dt, val_C6_qt)
    val_C_dt = sum_C_1_5_dt + val_C6_dt
    val_C_qt = sum_C_1_5_qt + val_C6_qt
    set_val('C', val_C_dt, val_C_qt)

    val_D_dt, val_D_qt = get_val('D')
    val_E1_dt = val_A1_dt + val_A2_dt + val_B7_dt + sum_C_1_5_dt + val_D_dt
    val_E1_qt = val_A1_qt + val_A2_qt + val_B7_qt + sum_C_1_5_qt + val_D_qt
    set_val('E.1', val_E1_dt, val_E1_qt)
    val_E2_dt = val_A3_dt + val_B8_dt + val_C6_dt
    val_E2_qt = val_A3_qt + val_B8_qt + val_C6_qt
    set_val('E.2', val_E2_dt, val_E2_qt)
    val_E_dt = val_E1_dt + val_E2_dt
    val_E_qt = val_E1_qt + val_E2_qt
    set_val('E', val_E_dt, val_E_qt)
    val_F_dt, val_F_qt = get_val('F')
    val_SCL_dt = val_E1_dt - val_F_dt
    val_SCL_qt = val_E1_qt - val_F_qt
    set_val('SCL', val_SCL_dt, val_SCL_qt)

    if changed[0]:
        st.session_state.sub_df = calculated_df
        st.rerun()

    st.divider()
    if st.button("💾 Lưu tất cả dữ liệu", type="primary", key=f"save_{wid_key}"):
        if not ten_ct.strip():
            st.error("Vui lòng nhập Tên Công trình (bắt buộc)!")
        else:
            try:
                # 0. Save chi tiết công trình (database mới)
                save_all_chitiet(sel_ma, pakt_data, kh_data, kq_data, hd_data, vt_data, nt_data)
                # 1. Update Tổng hợp.xlsx
                df_save = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Tổng hợp.xlsx'))
                for col in df_save.columns:
                    if 'Mã' in str(col) and 'công trình' in str(col).lower():
                        mask = df_save[col].astype(str).str.strip() == sel_ma
                        if mask.any():
                            ridx = mask.idxmax()
                            for c2 in df_save.columns:
                                if 'Trạng thái' in str(c2): df_save.at[ridx, c2] = new_tt
                                if 'Nội dung' in str(c2) and 'sửa chữa' in str(c2).lower(): df_save.at[ridx, c2] = new_noidung
                                if 'Tiến độ' in str(c2): df_save.at[ridx, c2] = new_tiendo
                        break
                df_save.to_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Tổng hợp.xlsx'), index=False)
                
                # 2. Update database_cong_trinh.xlsx
                main_row = {
                    'STT': stt, 'Tên Công trình': ten_ct, 'Mã CT': ma_ct, 'Kế hoạch': ke_hoach,
                    'Số Phương án': so_pa, 'Ngày Phương án': ngay_pa, 'Giá trị Phương án': gt_pa,
                    'Số Dự toán': so_dt, 'Ngày Dự toán': ngay_dt, 'Giá trị Dự toán': gt_dt,
                    'Số Hợp đồng thiết kế': so_hdtk, 'Ngày Hợp đồng thiết kế': ngay_hdtk, 'Giá trị Hợp đồng thiết kế': gt_hdtk,
                    'Số Hợp đồng giám sát': so_hdgs, 'Ngày Hợp đồng giám sát': ngay_hdgs, 'Giá trị Hợp đồng giám sát': gt_hdgs,
                    'Số Hợp đồng xây lắp': so_hdxl, 'Ngày Hợp đồng xây lắp': ngay_hdxl, 'Giá trị Hợp đồng xây lắp': gt_hdxl,
                    'Giá trị phát sinh': gt_ps, 'Giá trị VT thừa': gt_vtt, 'Giá trị VTTH': gt_vtth,
                    'Số Q.định phê duyệt QT công trình': so_qd, 'Ngày Q.định phê duyệt QT công trình': ngay_qd, 
                    'Giá trị Q.định phê duyệt QT công trình': gt_qd, 'Số tiền bằng chữ': so_tien_chu, 
                    'Ghi chú': ghi_chu, 'Đơn vị QL': don_vi_ql,
                    'Căn cứ pháp lý': can_cu_phap_ly, 'Khối lượng công việc': khoi_luong_cv,
                    'Ngày khởi công': ngay_khoi_cong, 'Ngày hoàn thành': ngay_hoan_thanh
                }
                rows_to_add = [main_row]
                for index, row in edited_sub_df.iterrows():
                    if pd.notna(row['Tên Hạng mục']) and str(row['Tên Hạng mục']).strip() != "":
                        sub_row = {col: None for col in ALL_COLUMNS}
                        sub_row['STT'] = row.get('STT')
                        sub_row['Tên Công trình'] = row.get('Tên Hạng mục')
                        sub_row['Mã CT'] = ma_ct
                        sub_row['Giá trị Dự toán'] = row.get('Giá trị Dự toán')
                        sub_row['Giá trị Q.định phê duyệt QT công trình'] = row.get('Giá trị quyết toán')
                        sub_row['Ghi chú'] = row.get('Chênh lệch')
                        rows_to_add.append(sub_row)
                
                new_data = pd.DataFrame(rows_to_add)
                db_df = load_db_data()
                
                # Delete old project records by checking Mã CT
                if ma_ct:
                    start_indices = db_df.index[db_df['Mã CT'].astype(str).str.strip() == ma_ct].tolist()
                    if not start_indices:
                        # Fallback to name match
                        start_indices = db_df.index[db_df['Tên Công trình'] == ten_ct].tolist()
                        
                    if start_indices:
                        start_idx = start_indices[0]
                        end_idx = len(db_df)
                        for i in range(start_idx + 1, len(db_df)):
                            val = str(db_df.at[i, 'STT']).strip().upper()
                            if val in ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']:
                                end_idx = i
                                break
                        db_df = db_df.drop(db_df.index[start_idx:end_idx])
                
                updated_df = pd.concat([db_df, new_data], ignore_index=True)
                updated_df.to_excel(DB_FILE, index=False)
                
                if 'sub_df' in st.session_state:
                    del st.session_state['sub_df']
                
                st.success("🎉 Lưu trữ dữ liệu thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi khi lưu dữ liệu: {e}")
