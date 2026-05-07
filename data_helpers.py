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
    """
    path = os.path.join(BASE_DIR, 'PM_092.xlsx')
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, header=None)
    
    result = {}
    current_ct = None
    
    for i, row in df.iterrows():
        cell0 = str(row[0]).strip()
        # Tìm dòng "Công trình: VTADxxxxxx - ..."
        if cell0.startswith('Công trình:'):
            match = re.search(r'(VTAD\d+)', cell0)
            if match:
                current_ct = match.group(1)
        # Tìm dòng "Tổng số dư cuối kỳ _ CÔNG TRÌNH"
        if 'Tổng số dư cuối kỳ _ CÔNG TRÌNH' in cell0 and current_ct:
            try:
                no_val = float(row[4]) if pd.notna(row[4]) else 0
            except:
                no_val = 0
            result[current_ct] = int(no_val)
            current_ct = None
    
    return result


def get_trang_thai_list():
    """Danh sách trạng thái từ file Tổng hợp"""
    return [
        'Đang thi công',
        'Lập PAKT-Tổng dự toán',
        'Lập kế hoạch đầu thầu',
        'Hoàn thành',
        'Nghiệm thu',
    ]


def get_nguon_list():
    """Danh sách nguồn"""
    return ['PM_092', 'Tổng Hợp']
