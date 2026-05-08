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
    if pd.isna(v) or v == "" or v is None: return "0"
    try: return f"{int(float(v)):,}"
    except: return "0"

def parse_num_val(s):
    if not s: return 0
    try: return int(str(s).replace(',', '').replace('.', '').strip())
    except: return 0
