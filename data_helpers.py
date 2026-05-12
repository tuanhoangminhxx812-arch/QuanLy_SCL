"""
Module hỗ trợ đọc dữ liệu từ file Tổng hợp.xlsx và PM_092.xlsx
"""
import pandas as pd
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def load_tonghop():
    """Đọc file Tổng hợp.xlsx - danh sách công trình chính"""
    path = os.path.join(BASE_DIR, 'Tổng hợp.xlsx')
    if not os.path.exists(path):
        return pd.DataFrame()
    df = pd.read_excel(path)
    # Chuẩn hóa tên cột
    col_map = {}
    for c in df.columns:
        cl = str(c).strip()
        if 'Mã' in cl and 'công trình' in cl.lower():
            col_map[c] = 'Mã CT'
        elif 'Tên công trình' in cl:
            col_map[c] = 'Tên công trình'
        elif 'Số QĐ' in cl or 'phê duyệt danh mục' in cl:
            col_map[c] = 'Số QĐ phê duyệt'
        elif 'Nội dung' in cl and ('sửa chữa' in cl.lower() or 'sữa chữa' in cl.lower() or 'scl' in cl.lower()):
            col_map[c] = 'Nội dung SCL'
        elif 'Tiến độ' in cl:
            col_map[c] = 'Tiến độ'
        elif cl == 'Năm':
            col_map[c] = 'Năm'
        elif 'khái toán' in cl.lower():
            col_map[c] = 'Khái toán'
        elif 'Tên hạng mục' in cl:
            col_map[c] = 'Tên hạng mục'
        elif 'Trạng thái' in cl:
            col_map[c] = 'Trạng thái'
        elif 'thực hiện' in cl.lower():
            col_map[c] = 'Thực hiện'
        elif 'quyết toán' in cl.lower():
            col_map[c] = 'Quyết toán'
    df = df.rename(columns=col_map)
    # Loại bỏ dòng tổng
    if 'STT' in df.columns:
        df = df[df['STT'].notna()].copy()
    return df


def load_pm092():
    """Đọc file PM_092.xlsx - lấy giá trị thực hiện theo mã công trình.
    Trả về dict: {mã_ct: tổng_nợ}
    
    Cấu trúc file PM_092:
    - "Công trình: VTADxxxxxx - ..."     -> bắt đầu 1 công trình
    -   "Hạng mục: ..."                  -> hạng mục con
    -     "Tài khoản: ..."               -> tài khoản kế toán
    -       các dòng giao dịch (INV)
    -       "Cộng phát sinh"             -> tổng phát sinh tài khoản
    -       "Số dư cuối kỳ"              -> số dư cuối kỳ tài khoản
    -   "Tổng số dư cuối kỳ _ HẠNG MỤC" -> tổng hạng mục (có thể gộp nhiều CT)
    -   "Tổng số dư cuối kỳ _ CÔNG TRÌNH"-> tổng công trình (có thể gộp nhiều CT)
    
    LƯU Ý: Cùng 1 mã CT có thể xuất hiện nhiều lần trong file (ở các phần
    tài khoản khác nhau). Dòng "Tổng số dư cuối kỳ _ CÔNG TRÌNH" có thể gộp
    chung nhiều mã CT khác nhau -> KHÔNG đáng tin cho từng CT riêng lẻ.
    
    => Giải pháp: Đọc "Số dư cuối kỳ" (không có hậu tố) ngay sau mỗi lần
    xuất hiện của 1 mã CT cụ thể, rồi cộng dồn.
    """
    path = os.path.join(BASE_DIR, 'PM_092.xlsx')
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, header=None)
    
    result = {}
    current_ct = None
    no_col_idx = 4  # Cột chứa giá trị Nợ (mặc định cột E = index 4)
    
    for i, row in df.iterrows():
        cell0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        # Tìm dòng "Công trình: VTADxxxxxx - ..."
        if cell0.startswith('Công trình:'):
            match = re.search(r'(VTAD\d+)', cell0)
            if match:
                current_ct = match.group(1)
            continue
        
        # Tìm dòng "Số dư cuối kỳ" (KHÔNG có hậu tố _ HẠNG MỤC hay _ CÔNG TRÌNH)
        # Đây là số dư cuối kỳ cho đúng phần tài khoản hiện tại của mã CT hiện tại
        if cell0 == 'Số dư cuối kỳ' and current_ct:
            try:
                no_val = float(row.iloc[no_col_idx]) if pd.notna(row.iloc[no_col_idx]) else 0
            except:
                no_val = 0
            result[current_ct] = result.get(current_ct, 0) + int(no_val)
            # Không reset current_ct ở đây, vì có thể còn tài khoản khác
            # của cùng 1 mã CT ngay phía dưới
            continue
        
        # Khi gặp dòng Tổng, reset current_ct
        if 'Tổng số dư cuối kỳ _ CÔNG TRÌNH' in cell0:
            current_ct = None
            continue
    
    return result


def get_trang_thai_list():
    """Danh sách trạng thái từ file Tổng hợp"""
    return [
        'Đang thi công',
        'Lập PAKT-Tổng dự toán',
        'Lập kế hoạch đấu thầu',
        'Hoàn thành',
        'Nghiệm thu',
    ]


def get_nguon_list():
    """Danh sách nguồn"""
    return ['PM_092', 'Tổng Hợp']


def load_gia_tri_hop_dong():
    """Đọc giá trị hợp đồng cho mỗi công trình.
    Lấy từ sheet kq_dau_thau trong chi_tiet_cong_trinh.xlsx.
    Ưu tiên cột 'Giá trị hợp đồng', fallback 'GT gói thầu trúng'.
    Cộng giá trị của gói XL + TB.
    Trả về dict: {Mã CT: tổng giá trị HĐ}
    """
    chitiet_path = os.path.join(BASE_DIR, 'chi_tiet_cong_trinh.xlsx')
    if not os.path.exists(chitiet_path):
        return {}
    try:
        df = pd.read_excel(chitiet_path, sheet_name='kq_dau_thau')
    except Exception:
        return {}
    if df.empty or 'Mã CT' not in df.columns:
        return {}
    
    result = {}
    for _, row in df.iterrows():
        ma = str(row.get('Mã CT', '')).strip()
        if not ma:
            continue
        # Ưu tiên Giá trị hợp đồng, fallback GT gói thầu trúng
        gt = row.get('Giá trị hợp đồng', 0)
        try:
            gt = int(float(gt)) if pd.notna(gt) else 0
        except:
            gt = 0
        if gt == 0:
            gt = row.get('GT gói thầu trúng', 0)
            try:
                gt = int(float(gt)) if pd.notna(gt) else 0
            except:
                gt = 0
        result[ma] = result.get(ma, 0) + gt
    return result

