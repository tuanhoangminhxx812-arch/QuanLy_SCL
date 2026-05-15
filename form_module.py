import streamlit as st
import pandas as pd
import os
import re

HO_SO_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Ho_so_cong_trinh')
DB_FILE = 'database_cong_trinh.xlsx'

DANH_MUC_HO_SO = [
    ('a', 'Biên bản khảo sát và đánh giá chi tiết hiện trạng VTTB, công trình'),
    ('b', 'Văn bản chấp thuận thông qua danh mục công trình SCL'),
    ('c', 'PAKT-DT công trình được phê duyệt'),
    ('d', 'Quyết định giao kế hoạch SCL'),
    ('đ', 'Kế hoạch lựa chọn nhà thầu, hồ sơ mời thầu, kết quả lựa chọn nhà thầu được phê duyệt'),
    ('e', 'Các hợp đồng, phụ lục hợp đồng (nếu có), các bảo lãnh theo qui định'),
    ('g', 'Hồ sơ hoàn công và Biên bản nghiệm thu'),
    ('h', 'Hồ sơ quyết toán công trình'),
    ('i', 'Quyết định phê duyệt quyết toán công trình'),
]

ALL_COLUMNS = [
    'STT', 'Tên Công trình', 'Mã CT', 'Kế hoạch', 'Số Phương án', 'Ngày Phương án', 
    'Giá trị Phương án', 'Số Dự toán', 'Ngày Dự toán', 'Giá trị Dự toán', 
    'Số Hợp đồng thiết kế', 'Ngày Hợp đồng thiết kế', 'Giá trị Hợp đồng thiết kế', 
    'Số Hợp đồng giám sát', 'Ngày Hợp đồng giám sát', 'Giá trị Hợp đồng giám sát', 
    'Số Hợp đồng xây lắp', 'Ngày Hợp đồng xây lắp', 'Giá trị Hợp đồng xây lắp', 
    'Giá trị phát sinh', 'Giá trị VT thừa', 'Giá trị VTTH', 
    'Số Q.định phê duyệt QT công trình', 'Ngày Q.định phê duyệt QT công trình', 
    'Giá trị Q.định phê duyệt QT công trình', 'Số tiền bằng chữ', 'Ghi chú', 'Đơn vị QL',
    'Căn cứ pháp lý', 'Khối lượng công việc', 'Ngày khởi công', 'Ngày hoàn thành'
]

def safe_folder_name(name):
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = name.strip('. ')
    if not name: name = 'unknown'
    return name[:100]

def get_project_folder(project_name):
    return os.path.join(HO_SO_BASE_DIR, safe_folder_name(project_name))

def get_category_folder(project_name, cat_key, cat_name):
    folder_name = f"{cat_key}) {safe_folder_name(cat_name)}"
    return os.path.join(get_project_folder(project_name), folder_name)

def get_safe_long_path(path):
    abs_path = os.path.abspath(path)
    if os.name == 'nt' and not abs_path.startswith('\\\\?\\'):
        return "\\\\?\\" + os.path.normpath(abs_path)
    return abs_path

def open_file_external(folder_path, file_name):
    """Mở file bằng ứng dụng mặc định, xử lý đường dẫn dài.
    Copy file ra thư mục temp (path ngắn) rồi mở từ đó.
    """
    import shutil, tempfile
    long_folder = get_safe_long_path(folder_path)
    src = os.path.join(long_folder, file_name)
    # Tạo thư mục temp ngắn gọn
    tmp_dir = os.path.join(tempfile.gettempdir(), 'scl_viewer')
    os.makedirs(tmp_dir, exist_ok=True)
    dst = os.path.join(tmp_dir, file_name)
    shutil.copy2(src, dst)
    os.startfile(dst)

def ensure_project_folders(project_name):
    project_dir = get_project_folder(project_name)
    # Add Windows long path prefix to avoid WinError 206 only on Windows
    long_project_dir = get_safe_long_path(project_dir)
    os.makedirs(long_project_dir, exist_ok=True)
    for cat_key, cat_name in DANH_MUC_HO_SO:
        cat_dir = get_category_folder(project_name, cat_key, cat_name)
        long_cat_dir = get_safe_long_path(cat_dir)
        os.makedirs(long_cat_dir, exist_ok=True)
    return project_dir

def list_files_in_folder(folder_path):
    long_folder_path = get_safe_long_path(folder_path)
    if not os.path.exists(long_folder_path): return []
    return [f for f in os.listdir(long_folder_path) if os.path.isfile(os.path.join(long_folder_path, f))]

def load_db_data():
    if os.path.exists(DB_FILE):
        df = pd.read_excel(DB_FILE)
        if 'Ghi chú' in df.columns:
            df['Ghi chú'] = df['Ghi chú'].astype(str).replace(['nan', 'None'], '')
        return df
    else:
        df = pd.DataFrame(columns=ALL_COLUMNS)
        df.to_excel(DB_FILE, index=False)
        return df

def doc_so_vn(n):
    if not n: return "Không đồng"
    try: n = int(n)
    except: return ""
    if n == 0: return "Không đồng"
    if n < 0: return "Âm " + doc_so_vn(-n)
    digits = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
    units = ["", "nghìn", "triệu", "tỷ"]
    def decode_3(num, full):
        res = []
        h = num // 100
        t = (num % 100) // 10
        u = num % 10
        if full or h > 0: res.append(f"{digits[h]} trăm")
        if t > 1:
            res.append(f"{digits[t]} mươi")
            if u == 1: res.append("mốt")
            elif u == 5: res.append("lăm")
            elif u > 0: res.append(digits[u])
        elif t == 1:
            res.append("mười")
            if u == 5: res.append("lăm")
            elif u > 0: res.append(digits[u])
        elif res and u > 0: res.append(f"lẻ {digits[u]}")
        elif not res and u > 0: res.append(digits[u])
        return " ".join(res)
    blocks = []
    while n > 0:
        blocks.append(n % 1000)
        n //= 1000
    words = []
    for i, block in enumerate(blocks):
        if block == 0 and i > 0: continue
        full = (i < len(blocks)-1)
        unit = units[i % 4]
        for _ in range(i // 4): unit += " tỷ" 
        words.append(f"{decode_3(block, full)} {unit}".strip())
    res = " ".join(reversed(words)).strip()
    res = res.replace("  ", " ").strip()
    return res[0].upper() + res[1:] + " đồng"

def format_num_val(v):
    """Format số tiền kiểu Việt Nam: dấu . phân cách hàng nghìn, dấu , phân cách thập phân."""
    if pd.isna(v) or v == "" or v is None: return "0"
    try:
        f_val = float(v)
        if f_val == int(f_val):
            # Số nguyên: 13254 → "13.254"
            return f"{int(f_val):,}".replace(",", ".")
        else:
            # Số thập phân: 13254.6 → "13.254,6"
            int_part = int(f_val)
            dec_part = round(f_val - int_part, 4)
            dec_str = f"{dec_part:.10f}".rstrip('0').lstrip('0').lstrip('.')
            int_formatted = f"{int_part:,}".replace(",", ".")
            return f"{int_formatted},{dec_str}"
    except: return "0"

def parse_num_val(s):
    """Parse số tiền kiểu VN (dấu . hàng nghìn, dấu , thập phân) → số."""
    if not s: return 0
    s = str(s).strip()
    try:
        # Kiểu VN: "13.254,6" → 13254.6
        if ',' in s and '.' in s:
            s = s.replace('.', '').replace(',', '.')
        elif ',' in s and '.' not in s:
            # Có thể là "13,254" (VN decimal) hoặc legacy "13,254" (EN thousands)
            # Nếu chỉ 1 dấu , và phần sau , < 3 ký tự → thập phân VN
            parts = s.split(',')
            if len(parts) == 2 and len(parts[1]) <= 3:
                s = s.replace(',', '.')
            else:
                s = s.replace(',', '')
        elif '.' in s:
            # "13.254" — kiểu VN hàng nghìn → bỏ dấu .
            parts = s.split('.')
            if all(len(p) <= 3 for p in parts[1:]):
                s = s.replace('.', '')
        
        val = float(s)
        return int(val) if val == int(val) else val
    except: return 0

# ============================================================
# CHI TIẾT CÔNG TRÌNH — Database mới
# ============================================================

CHITIET_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'chi_tiet_cong_trinh.xlsx')

LOAI_HOP_DONG = ['Xây lắp', 'Thiết bị', 'Tư vấn thiết kế', 'Tư vấn giám sát', 'Khác']
HINH_THUC_HD = ['Trọn gói', 'Hợp đồng theo đơn giá cố định', 'Hợp đồng theo đơn giá điều chỉnh', 'Hợp đồng theo thời gian', 'Khác']

# Column definitions for each sheet
PAKT_DT_COLS = ['Mã CT', 'Số QĐ phê duyệt', 'Ngày phê duyệt', 'Giá trị dự toán']
KH_DAU_THAU_COLS = ['Mã CT', 'Loại gói', 'Số QĐ phê duyệt KH', 'Ngày phê duyệt', 'GT gói thầu']
KQ_DAU_THAU_COLS = ['Mã CT', 'Loại gói', 'Số QĐ phê duyệt KQ', 'Ngày phê duyệt', 'GT gói thầu trúng', 'Số hợp đồng', 'Giá trị hợp đồng']
HOP_DONG_COLS = [
    'Mã CT', 'Loại HĐ', 'Gói thầu', 'Tên nhà thầu', 'Hình thức HĐ',
    'Số hợp đồng', 'Ngày ký HĐ', 'Ngày hiệu lực', 'Tên hợp đồng',
    'Giá trị HĐ', 'Giá trị bảo lãnh', 'Thời gian thực hiện',
    'Giá trị thực hiện HĐ', 'Số BB nghiệm thu', 'Ngày nghiệm thu'
]
VAT_TU_COLS = ['Mã CT', 'TCty cấp', 'ĐV cấp']
NGHIEM_THU_QT_COLS = ['Mã CT', 'Ngày nghiệm thu CT', 'Giá trị quyết toán CT', 'Ghi chú']

SHEET_COLS_MAP = {
    'pakt_dt': PAKT_DT_COLS,
    'kh_dau_thau': KH_DAU_THAU_COLS,
    'kq_dau_thau': KQ_DAU_THAU_COLS,
    'hop_dong': HOP_DONG_COLS,
    'vat_tu': VAT_TU_COLS,
    'nghiem_thu_qt': NGHIEM_THU_QT_COLS,
}

def _ensure_chitiet_file():
    """Tạo file chi_tiet_cong_trinh.xlsx nếu chưa tồn tại."""
    if not os.path.exists(CHITIET_FILE):
        with pd.ExcelWriter(CHITIET_FILE, engine='openpyxl') as writer:
            for sheet_name, cols in SHEET_COLS_MAP.items():
                pd.DataFrame(columns=cols).to_excel(writer, sheet_name=sheet_name, index=False)

def load_chitiet_sheet(sheet_name):
    """Đọc 1 sheet từ chi_tiet_cong_trinh.xlsx."""
    _ensure_chitiet_file()
    try:
        df = pd.read_excel(CHITIET_FILE, sheet_name=sheet_name)
        return df
    except Exception:
        return pd.DataFrame(columns=SHEET_COLS_MAP.get(sheet_name, []))

def save_chitiet_sheet(sheet_name, df):
    """Ghi 1 sheet vào chi_tiet_cong_trinh.xlsx (giữ nguyên các sheet khác)."""
    _ensure_chitiet_file()
    # Read all existing sheets
    all_sheets = {}
    try:
        with pd.ExcelFile(CHITIET_FILE) as xls:
            for s in xls.sheet_names:
                all_sheets[s] = pd.read_excel(xls, sheet_name=s)
    except Exception:
        pass
    # Update target sheet
    all_sheets[sheet_name] = df
    # Ensure all sheets exist
    for s, cols in SHEET_COLS_MAP.items():
        if s not in all_sheets:
            all_sheets[s] = pd.DataFrame(columns=cols)
    # Write
    with pd.ExcelWriter(CHITIET_FILE, engine='openpyxl') as writer:
        for s_name in SHEET_COLS_MAP.keys():
            if s_name in all_sheets:
                all_sheets[s_name].to_excel(writer, sheet_name=s_name, index=False)

def load_chitiet_by_ma(sheet_name, ma_ct):
    """Đọc dữ liệu theo Mã CT từ 1 sheet."""
    df = load_chitiet_sheet(sheet_name)
    if df.empty or 'Mã CT' not in df.columns:
        return pd.DataFrame(columns=SHEET_COLS_MAP.get(sheet_name, []))
    return df[df['Mã CT'].astype(str).str.strip() == str(ma_ct).strip()].copy()

def save_chitiet_by_ma(sheet_name, ma_ct, new_data_df):
    """Ghi dữ liệu cho 1 Mã CT vào sheet (xóa dữ liệu cũ của CT đó, thay bằng mới)."""
    df = load_chitiet_sheet(sheet_name)
    # Remove old records for this project
    if not df.empty and 'Mã CT' in df.columns:
        df = df[df['Mã CT'].astype(str).str.strip() != str(ma_ct).strip()]
    # Append new data
    if not new_data_df.empty:
        df = pd.concat([df, new_data_df], ignore_index=True)
    save_chitiet_sheet(sheet_name, df)

def load_hopdong_list(ma_ct):
    """Đọc danh sách hợp đồng của 1 CT."""
    return load_chitiet_by_ma('hop_dong', ma_ct)

def save_hopdong_list(ma_ct, hd_df):
    """Lưu danh sách hợp đồng của 1 CT."""
    save_chitiet_by_ma('hop_dong', ma_ct, hd_df)
