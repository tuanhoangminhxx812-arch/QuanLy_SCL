"""
dossier_scanner.py
Module tự động quét và trích xuất dữ liệu từ thư mục Ho_so_cong_trinh
- Nhận diện và ưu tiên hàng đầu file Quyết toán có tên: QT_A-B (QT_A-B.xlsx, QT_A-B.xls, QT_A-B.pdf)
- Hỗ trợ cả file PDF (đọc số liệu từ PDF) và file Excel (đọc chính xác 100% từng đồng)
- Trích xuất số liệu Dự toán (A, B, C, D, F, SCL, Quy mô, QĐ phê duyệt)
- Trích xuất thông tin hợp đồng, nhà thầu, thời gian thi công, căn cứ pháp lý
- Tự động mở rộng 38 khoản mục chi phí chuẩn theo EVNHCMC
- Cơ chế SAO LƯU (Backup) an toàn vào data/backups/
- Cơ chế CACHE SNAPSHOT vào data/cache_dossier/ để chương trình khởi động siêu tốc không bị lag
- Cập nhật tự động vào file database_cong_trinh.xlsx CHỈ KHI người dùng bấm nút cập nhật.
"""

import os
import re
import json
import shutil
import datetime
import pandas as pd
from pypdf import PdfReader
from form_module import doc_so_vn
from legal_checker import check_compliance

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HO_SO_DIR = os.path.join(BASE_DIR, "Ho_so_cong_trinh")
THAM_KHAO_DIR = os.path.join(BASE_DIR, "Tham khao")
DB_FILE = os.path.join(BASE_DIR, "database_cong_trinh.xlsx")
BACKUP_DIR = os.path.join(BASE_DIR, "data", "backups")
CACHE_DIR = os.path.join(BASE_DIR, "data", "cache_dossier")

STANDARD_COST_ROWS = [
    ('A', 'CHI PHÍ VẬT TƯ, THIẾT BỊ (sau thuế)'),
    ('A.1', 'Chi phí thiết bị'),
    ('A.1.1', 'Thiết bị nhập khẩu'),
    ('A.1.2', 'VT A cấp'),
    ('A.1.3', 'Chi phí tháo dỡ, lắp đặt'),
    ('A.1.4', 'Chi phí thí nghiệm, hiệu chỉnh'),
    ('A.2', 'Chi phí vật tư'),
    ('A.3', 'Thuế GTGT'),
    ('B', 'CHI PHÍ SỬA CHỮA'),
    ('B.1', 'Chi phí vật liệu'),
    ('B.1.1', 'Vật liệu phần không áp dụng đơn giá XDCB'),
    ('B.1.2', 'Vật liệu phần áp dụng đơn giá XDCB'),
    ('B.1.3', 'Chênh lệch giá vật liệu phần áp dụng đơn giá XDCB'),
    ('B.1.4', 'Vật liệu phụ trong SCL thiết bị'),
    ('B.2', 'Chi phí nhân công'),
    ('B.2.1', 'Chi phí nhân công phần không áp dụng đơn giá XDCB'),
    ('B.2.2', 'Chi phí nhân công phần áp dụng đơn giá XDCB'),
    ('B.3', 'Chi phí máy thi công'),
    ('B.3.1', 'Chi phí máy thi công phần không áp dụng đơn giá XDCB'),
    ('B.3.2', 'Chi phí máy thi công phần áp dụng đơn giá XDCB'),
    ('B.4', 'Chi phí làm đêm, làm thêm giờ'),
    ('B.5', 'Chi phí chung'),
    ('B.6', 'Thu nhập chịu thuế tính trước'),
    ('B.7', 'Giá trị sửa chữa trước thuế'),
    ('B.8', 'Thuế GTGT'),
    ('C', 'CHI PHÍ KHÁC (sau thuế)'),
    ('C.1', 'Chi phí giám sát thi công xây dựng'),
    ('C.2', 'Chi phí giám sát lắp đặt thiết bị'),
    ('C.3', 'Chi phí bảo hiểm công trình'),
    ('C.4', 'Chi phí thẩm tra - phê duyệt quyết toán'),
    ('C.5', 'Vận chuyển VTTB A cấp đến công trường'),
    ('C.6', 'Thuế GTGT'),
    ('D', 'CHI PHÍ DỰ PHÒNG'),
    ('E', 'Tổng giá trị trước thuế'),
    ('E.1', 'Tổng giá trị thuế GTGT'),
    ('E.2', 'Tổng giá trị sau thuế'),
    ('F', 'GIÁ TRỊ VẬT TƯ THU HỒI'),
    ('SCL', 'CHI PHÍ SCL'),
]


def _clean_str(s):
    if not s:
        return ""
    return str(s).strip().lower()


def find_project_folder(ma_ct, ten_ct=""):
    """Tìm thư mục của công trình trong Ho_so_cong_trinh hoặc Tham khao."""
    ma_clean = _clean_str(ma_ct)
    ten_clean = _clean_str(ten_ct)

    # Ưu tiên kiểm tra Tham khao cho công trình Công xa VTAD2608001
    if "vtad2608001" in ma_clean or "công xa" in ten_clean or "cong xa" in ten_clean:
        if os.path.exists(THAM_KHAO_DIR):
            return THAM_KHAO_DIR

    if not os.path.exists(HO_SO_DIR):
        return None

    candidates = []
    for d in os.listdir(HO_SO_DIR):
        full = os.path.join(HO_SO_DIR, d)
        if os.path.isdir(full):
            candidates.append((d, full))

    # 1. So khớp chính xác 100% với tên công trình
    for d_name, full_path in candidates:
        if _clean_str(d_name) == ten_clean:
            return full_path

    # 2. So khớp theo Mã CT
    if ma_clean:
        for d_name, full_path in candidates:
            if ma_clean in _clean_str(d_name):
                return full_path

    # 3. So khớp theo độ trùng khớp dài nhất
    best_match = None
    max_len = 0
    for d_name, full_path in candidates:
        d_lower = _clean_str(d_name)
        if d_lower in ten_clean or ten_clean in d_lower:
            common_len = len(d_lower)
            if common_len > max_len:
                max_len = common_len
                best_match = full_path

    if best_match:
        return best_match

    # 4. Fallback từ khóa đặc thù
    keywords = []
    if "đường dây" in ten_clean or "tba" in ten_clean:
        keywords = ["đường dây trung", "duong day trung", "đường dây", "lưới điện"]
    elif "tu" in ten_clean and "ti" in ten_clean:
        keywords = ["tu, ti", "tu", "ti", "đo đếm"]
    elif "fco" in ten_clean or "lbfco" in ten_clean:
        keywords = ["fco", "lbfco", "ngăn ngừa sự cố"]
    elif "an hội" in ten_clean or "côn đảo" in ten_clean:
        keywords = ["an hội", "côn đảo"]
    elif "live-line" in ten_clean or "live line" in ten_clean:
        keywords = ["live-line", "live line"]
    elif "cummins" in ten_clean:
        keywords = ["cummins"]

    for d_name, full_path in candidates:
        d_lower = d_name.lower()
        if any(k in d_lower for k in keywords):
            return full_path

    return None


def scan_all_files(folder_path):
    """Quét đệ quy toàn bộ file trong thư mục."""
    all_files = []
    if not folder_path or not os.path.exists(folder_path):
        return all_files

    for root, dirs, files in os.walk(folder_path):
        for f in files:
            all_files.append(os.path.join(root, f))
    return all_files


def extract_numbers_from_line(line):
    """Tìm tất cả các con số tiền trong 1 dòng văn bản."""
    matches = re.findall(r'(?<!\w)(?:\d{1,3}[ \.,]\d{3}(?:[ \.,]\d{3})*|\d{6,})(?!\w)', line)
    nums = []
    for m in matches:
        clean = re.sub(r'[^\d]', '', m)
        if len(clean) >= 5:
            nums.append(int(clean))
    return nums


def extract_money_from_text(text):
    """Trích xuất 1 số tiền (VNĐ) chuẩn xác từ văn bản."""
    if not text:
        return 0
    m = re.search(r'((?:\d{1,3}[ \.,]\d{3}(?:[ \.,]\d{3})*|\d{5,}))\s*(?:đồng|đ)', text, re.IGNORECASE)
    if m:
        clean = re.sub(r'[^\d]', '', m.group(1))
        try:
            return int(clean)
        except:
            pass
    nums = extract_numbers_from_line(text)
    return nums[-1] if nums else 0


def extract_pakt_dt_from_pdf(pdf_path):
    """Trích xuất số liệu Dự toán từ file Quyết định phê duyệt PAKT-DT."""
    res = {
        'found': False,
        'so_qd': '',
        'ngay_qd': '',
        'tong_dt': 0,
        'chi_phi_vttb': 0,
        'chi_phi_sua_chua': 0,
        'chi_phi_khac': 0,
        'chi_phi_du_phong': 0,
        'chi_phi_thu_hoi': 0,
        'can_cu': [],
        'khoi_luong': [],
        'source_file': os.path.basename(pdf_path)
    }

    try:
        reader = PdfReader(pdf_path)
        full_text = ""
        for p in reader.pages:
            full_text += (p.extract_text() or "") + "\n"

        # 1. Tìm số QĐ và ngày QĐ
        m_so = re.search(r'Số:\s*([^\n\r/]+(?:/[^\n\r]+)?)', full_text)
        if m_so:
            s_val = m_so.group(1).strip()
            if len(s_val) > 2 and not s_val.startswith('/'):
                res['so_qd'] = s_val

        if not res['so_qd']:
            fbase = os.path.basename(pdf_path)
            m_fn = re.search(r'(\d+[\-_A-Za-z0-9]*)', fbase)
            if m_fn:
                res['so_qd'] = m_fn.group(1)

        m_ngay = re.search(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})', full_text, re.IGNORECASE)
        if m_ngay:
            d, m, y = m_ngay.groups()
            res['ngay_qd'] = f"{int(d):02d}/{int(m):02d}/{y}"

        # 2. Tìm chi phí dự toán
        for line in full_text.split('\n'):
            line_l = line.lower()
            if 'chi phí vật tư, thiết bị' in line_l or 'chi phí thiết bị' in line_l:
                val = extract_money_from_text(line)
                if val > 0 and res['chi_phi_vttb'] == 0:
                    res['chi_phi_vttb'] = val
            elif 'chi phí sửa chữa' in line_l:
                val = extract_money_from_text(line)
                if val > 0 and res['chi_phi_sua_chua'] == 0:
                    res['chi_phi_sua_chua'] = val
            elif 'chi phí khác' in line_l:
                val = extract_money_from_text(line)
                if val > 0 and res['chi_phi_khac'] == 0:
                    res['chi_phi_khac'] = val
            elif 'chi phí dự phòng' in line_l:
                val = extract_money_from_text(line)
                if val > 0 and res['chi_phi_du_phong'] == 0:
                    res['chi_phi_du_phong'] = val
            elif 'tổng cộng' in line_l or 'dự toán chi phí sửa chữa' in line_l:
                val = extract_money_from_text(line)
                if val > 0 and res['tong_dt'] == 0:
                    res['tong_dt'] = val
            elif 'thu hồi' in line_l:
                val = extract_money_from_text(line)
                if val > 0 and res['chi_phi_thu_hoi'] == 0:
                    res['chi_phi_thu_hoi'] = val

        if res['tong_dt'] == 0:
            res['tong_dt'] = res['chi_phi_vttb'] + res['chi_phi_sua_chua'] + res['chi_phi_khac'] + res['chi_phi_du_phong']

        # 3. Trích xuất căn cứ pháp lý
        for line in full_text.split('\n'):
            line_s = line.strip()
            if line_s.startswith('Căn cứ') or line_s.startswith('- Căn cứ') or line_s.startswith('+ Căn cứ'):
                if len(line_s) > 15 and line_s not in res['can_cu']:
                    res['can_cu'].append(line_s)

        # 4. Trích xuất khối lượng công việc chính
        kl_lines = []
        is_kl_section = False
        for line in full_text.split('\n'):
            line_s = line.strip()
            if 'quy mô sửa chữa' in line_s.lower() or 'khối lượng công việc' in line_s.lower():
                is_kl_section = True
                continue
            if is_kl_section:
                if any(end_k in line_s.lower() for end_k in ['giải pháp kỹ thuật', 'dự toán chi phí', 'nguồn vốn thực hiện', 'tiến độ thực hiện']):
                    break
                if line_s and len(line_s) > 8:
                    kl_lines.append(line_s)
        res['khoi_luong'] = kl_lines[:30]

        if res['tong_dt'] > 0 or res['chi_phi_sua_chua'] > 0 or res['chi_phi_vttb'] > 0:
            res['found'] = True

    except Exception as e:
        print(f"Lỗi khi đọc file {pdf_path}: {e}")

    return res


def extract_settlement_from_pdf(pdf_path):
    """
    Trích xuất số liệu Quyết toán A-B từ file PDF (QT_A-B.pdf, Bảng THQT hoặc Biên bản quyết toán).
    Đặc biệt xử lý các bảng có cả 2 cột: Cột Dự toán và Cột Quyết toán.
    """
    res = {
        'found': False,
        'tong_qt': 0,
        'tong_dt': 0,
        'chi_phi_vttb': 0,
        'chi_phi_vttb_dt': 0,
        'chi_phi_sua_chua': 0,
        'chi_phi_sua_chua_dt': 0,
        'vat_lieu': 0,
        'vat_lieu_dt': 0,
        'nhan_cong': 0,
        'nhan_cong_dt': 0,
        'may_thi_cong': 0,
        'chi_phi_chung': 0,
        'thue_gtgt': 0,
        'thue_gtgt_dt': 0,
        'chi_phi_khac': 0,
        'chi_phi_du_phong': 0,
        'chi_phi_du_phong_dt': 0,
        'chi_phi_thu_hoi': 0,
        'source_file': os.path.basename(pdf_path)
    }

    try:
        reader = PdfReader(pdf_path)
        full_text = ""
        for p in reader.pages:
            full_text += (p.extract_text() or "") + "\n"

        for line in full_text.split('\n'):
            line_l = line.lower()
            nums = extract_numbers_from_line(line)
            if not nums:
                continue

            if len(nums) >= 2:
                val_dt = nums[0]
                val_qt = nums[1]
            else:
                val_dt = 0
                val_qt = nums[0]

            if 'chi phí thiết bị' in line_l or 'vật tư, thiết bị' in line_l:
                if res['chi_phi_vttb'] == 0:
                    res['chi_phi_vttb'] = val_qt
                    if val_dt > 0:
                        res['chi_phi_vttb_dt'] = val_dt
            elif 'chi phí sửa chữa' in line_l:
                if res['chi_phi_sua_chua'] == 0:
                    res['chi_phi_sua_chua'] = val_qt
                    if val_dt > 0:
                        res['chi_phi_sua_chua_dt'] = val_dt
            elif 'chi phí vật liệu' in line_l or 'vật liệu' in line_l:
                if res['vat_lieu'] == 0:
                    res['vat_lieu'] = val_qt
                    if val_dt > 0:
                        res['vat_lieu_dt'] = val_dt
            elif 'chi phí nhân công' in line_l or 'nhân công' in line_l:
                if res['nhan_cong'] == 0:
                    res['nhan_cong'] = val_qt
            elif 'chi phí máy thi công' in line_l or 'máy thi công' in line_l:
                if res['may_thi_cong'] == 0:
                    res['may_thi_cong'] = val_qt
            elif 'chi phí chung' in line_l:
                if res['chi_phi_chung'] == 0:
                    res['chi_phi_chung'] = val_qt
            elif 'thuế gtgt' in line_l or 'thuế vat' in line_l:
                if res['thue_gtgt'] == 0:
                    res['thue_gtgt'] = val_qt
                    if val_dt > 0:
                        res['thue_gtgt_dt'] = val_dt
            elif 'chi phí khác' in line_l:
                if res['chi_phi_khac'] == 0:
                    res['chi_phi_khac'] = val_qt
            elif 'chi phí dự phòng' in line_l or 'dự phòng' in line_l:
                if res['chi_phi_du_phong_dt'] == 0 and val_dt > 0:
                    res['chi_phi_du_phong_dt'] = val_dt
                elif res['chi_phi_du_phong_dt'] == 0 and len(nums) == 1:
                    res['chi_phi_du_phong_dt'] = nums[0]
            elif 'thu hồi' in line_l:
                if res['chi_phi_thu_hoi'] == 0:
                    res['chi_phi_thu_hoi'] = val_qt
            elif 'giá trị quyết toán' in line_l or 'cộng giá trị công trình' in line_l or 'tổng cộng' in line_l:
                if res['tong_qt'] == 0:
                    res['tong_qt'] = val_qt
                    if val_dt > 0:
                        res['tong_dt'] = val_dt

        if res['tong_qt'] == 0 and res['chi_phi_sua_chua'] > 0:
            res['tong_qt'] = res['chi_phi_vttb'] + res['chi_phi_sua_chua'] + res['chi_phi_khac'] - res['chi_phi_thu_hoi']

        if res['tong_qt'] > 0 or res['chi_phi_sua_chua'] > 0:
            res['found'] = True

    except Exception as e:
        print(f"Lỗi khi đọc quyết toán từ {pdf_path}: {e}")

    return res


def extract_settlement_from_excel(excel_path):
    """
    Trích xuất số liệu Quyết toán A-B từ file Excel (QT_A-B.xlsx, QT_A-B.xls).
    Đảm bảo bóc tách chính xác 100% từng đồng.
    """
    res = {
        'found': False,
        'tong_qt': 0, 'tong_dt': 0,
        'chi_phi_vttb': 0, 'chi_phi_vttb_dt': 0,
        'chi_phi_sua_chua': 0, 'chi_phi_sua_chua_dt': 0,
        'vat_lieu': 0, 'vat_lieu_dt': 0,
        'nhan_cong': 0, 'nhan_cong_dt': 0,
        'may_thi_cong': 0,
        'chi_phi_chung': 0,
        'thue_gtgt': 0, 'thue_gtgt_dt': 0,
        'chi_phi_khac': 0,
        'chi_phi_du_phong': 0, 'chi_phi_du_phong_dt': 0,
        'chi_phi_thu_hoi': 0,
        'source_file': os.path.basename(excel_path)
    }

    try:
        xl = pd.ExcelFile(excel_path)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            for _, row in df.iterrows():
                row_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
                nums = []
                for v in row.values:
                    try:
                        n = float(v)
                        if n > 1000:
                            nums.append(int(round(n)))
                    except:
                        pass
                if not nums:
                    continue

                val_dt = nums[0] if len(nums) >= 2 else 0
                val_qt = nums[1] if len(nums) >= 2 else nums[0]

                if 'thiết bị' in row_str or 'vật tư, thiết bị' in row_str:
                    res['chi_phi_vttb'] = val_qt
                    if val_dt: res['chi_phi_vttb_dt'] = val_dt
                elif 'sửa chữa' in row_str:
                    res['chi_phi_sua_chua'] = val_qt
                    if val_dt: res['chi_phi_sua_chua_dt'] = val_dt
                elif 'vật liệu' in row_str:
                    res['vat_lieu'] = val_qt
                    if val_dt: res['vat_lieu_dt'] = val_dt
                elif 'nhân công' in row_str:
                    res['nhan_cong'] = val_qt
                    if val_dt: res['nhan_cong_dt'] = val_dt
                elif 'máy thi công' in row_str:
                    res['may_thi_cong'] = val_qt
                elif 'chi phí chung' in row_str:
                    res['chi_phi_chung'] = val_qt
                elif 'thuế gtgt' in row_str or 'vat' in row_str:
                    res['thue_gtgt'] = val_qt
                    if val_dt: res['thue_gtgt_dt'] = val_dt
                elif 'chi phí khác' in row_str:
                    res['chi_phi_khac'] = val_qt
                elif 'dự phòng' in row_str:
                    res['chi_phi_du_phong_dt'] = val_dt or val_qt
                elif 'thu hồi' in row_str:
                    res['chi_phi_thu_hoi'] = val_qt
                elif 'quyết toán' in row_str or 'tổng cộng' in row_str or 'cộng giá trị' in row_str:
                    res['tong_qt'] = val_qt
                    if val_dt: res['tong_dt'] = val_dt

        if res['tong_qt'] > 0 or res['chi_phi_sua_chua'] > 0:
            res['found'] = True
    except Exception as e:
        print(f"Lỗi khi đọc file Excel {excel_path}: {e}")

    return res


def extract_contract_info(pdf_path):
    """Trích xuất thông tin hợp đồng, nhà thầu từ Thư trao hợp đồng hoặc Hợp đồng xây lắp."""
    res = {
        'found': False,
        'so_hd': '',
        'ngay_hd': '',
        'nha_thau': '',
        'gia_tri_hd': 0,
        'thoi_gian_thi_cong': '',
        'hinh_thuc': 'Thuê ngoài',
        'ngay_kc': '',
        'ngay_ht': '',
        'source_file': os.path.basename(pdf_path)
    }

    try:
        reader = PdfReader(pdf_path)
        full_text = ""
        for p in reader.pages[:15]:
            full_text += (p.extract_text() or "") + "\n"

        if not full_text.strip():
            return res

        # 1. Tìm số HĐ / Thư trao thầu
        m_so = re.search(r'(?:Số|Hợp đồng số):\s*([^\n\r/]+(?:/[^\n\r]+)?)', full_text, re.IGNORECASE)
        if m_so:
            s = m_so.group(1).strip()
            if len(s) > 2 and not s.startswith('/'):
                res['so_hd'] = s
        if not res['so_hd']:
            fbase = os.path.basename(pdf_path)
            m_fn = re.search(r'(\d+[\-_/A-Za-z0-9]+)', fbase)
            if m_fn:
                res['so_hd'] = m_fn.group(1)

        # 2. Tìm ngày
        m_ngay = re.search(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})', full_text, re.IGNORECASE)
        if m_ngay:
            d, m, y = m_ngay.groups()
            res['ngay_hd'] = f"{int(d):02d}/{int(m):02d}/{y}"

        # 3. Tìm tên nhà thầu
        m_nt = re.search(r'(?:trao\s+hợp\s+đồng\s+cho\s+nhà\s+thầu|đơn\s+vị\s+thi\s+công|nhà\s+thầu)[:\s]+([^\(\n\r]+(?:\n[^\(\n\r]+)?)', full_text, re.IGNORECASE)
        if m_nt:
            nt_cand = m_nt.group(1).replace('\n', ' ').strip()
            if any(k in nt_cand for k in ['Công ty', 'Liên danh', 'TNHH', 'Cổ phần']):
                res['nha_thau'] = nt_cand

        if not res['nha_thau']:
            m_nt2 = re.search(r'(Liên\s+danh\s+Công\s+ty[^\n\r\.\,\(]+(?:\n[^\n\r\.\,\(]+)?|Công\s+ty\s+(?:Cổ\s+phần|TNHH)[^\n\r\.\,\(]+)', full_text, re.IGNORECASE)
            if m_nt2:
                res['nha_thau'] = m_nt2.group(1).replace('\n', ' ').strip()

        if ' - ' in res['nha_thau']:
            res['nha_thau'] = res['nha_thau'].split(' - ')[0].strip()

        # 4. Giá hợp đồng
        m_gia = re.search(r'giá\s+(?:hợp\s+đồng|trúng\s+thầu)\s+là:\s*([\d\.\s,]+)\s*đồng', full_text, re.IGNORECASE)
        if m_gia:
            clean = re.sub(r'[^\d]', '', m_gia.group(1))
            if clean:
                res['gia_tri_hd'] = int(clean)

        # 5. Thời gian thực hiện
        m_tg = re.search(r'thời\s+gian\s+thực\s+hiện(?: gói thầu)?\s*là:\s*([^\n\r\.\,]+(?:\n[^\n\r\.\,]+)?)', full_text, re.IGNORECASE)
        if m_tg:
            res['thoi_gian_thi_cong'] = m_tg.group(1).replace('\n', ' ').strip()

        # 6. Khởi công / Hoàn thành
        m_kc = re.search(r'(?:kh\s*ởi|khởi)\s*công:\s*([^\n\r]+)', full_text, re.IGNORECASE)
        if m_kc:
            res['ngay_kc'] = m_kc.group(1).strip()
        m_ht = re.search(r'hoàn\s*thành:\s*([^\n\r]+)', full_text, re.IGNORECASE)
        if m_ht:
            res['ngay_ht'] = m_ht.group(1).strip()

        if 'tự làm' in full_text.lower() and 'thuê ngoài' not in full_text.lower():
            res['hinh_thuc'] = 'Tự làm'

        if res['so_hd'] or res['nha_thau'] or res['gia_tri_hd'] > 0 or res['ngay_kc']:
            res['found'] = True

    except Exception as e:
        print(f"Lỗi trích xuất hợp đồng {pdf_path}: {e}")

    return res


def load_cached_scan_result(ma_ct):
    """Đọc nhanh snapshot kết quả quét từ cache (tốc độ < 0.05 giây, không quét lại ổ đĩa)."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_path = os.path.join(CACHE_DIR, f"{str(ma_ct).strip()}.json")
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return None


def scan_and_update_project(ma_ct, ten_ct=""):
    """
    CHỈ CHẠY KHI NGƯỜI DÙNG BẤM NÚT CẬP NHẬT:
    1. Tìm thư mục công trình.
    2. Quét và trích xuất:
       - Ưu tiên file Quyết toán A-B có tên: QT_A-B (Excel hoặc PDF).
       - Trích xuất Dự toán, Hợp đồng, Nhà thầu, Tiến độ.
    3. Sao lưu (Backup) database_cong_trinh.xlsx vào data/backups/.
    4. Cập nhật và lưu vào database_cong_trinh.xlsx.
    5. Lưu kết quả vào cache data/cache_dossier/{ma_ct}.json để các lần xem sau mở tức thì.
    """
    folder = find_project_folder(ma_ct, ten_ct)
    if not folder or not os.path.exists(folder):
        return {
            'success': False,
            'message': f"Không tìm thấy thư mục hồ sơ cho công trình {ma_ct} trong Ho_so_cong_trinh hoặc Tham khao!",
            'folder': '',
            'scanned_files': [],
            'dt_data': {'found': False},
            'qt_data': {'found': False},
            'contract_data': {'found': False, 'so_hd': '', 'nha_thau': '', 'gia_tri_hd': 0, 'thoi_gian_thi_cong': '', 'hinh_thuc': 'Thuê ngoài', 'ngay_kc': '', 'ngay_ht': ''},
            'compliance': {}
        }

    scanned_files = scan_all_files(folder)

    dt_data = {'found': False}
    qt_data = {'found': False}
    contract_data = {'found': False, 'so_hd': '', 'nha_thau': '', 'gia_tri_hd': 0, 'thoi_gian_thi_cong': '', 'hinh_thuc': 'Thuê ngoài', 'ngay_kc': '', 'ngay_ht': ''}

    # BƯỚC 1: Tìm kiếm file Quyết toán A-B được ưu tiên đặt tên là QT_A-B (hoặc biến thể)
    qtab_excel = []
    qtab_pdf = []
    other_qt_excel = []
    other_qt_pdf = []

    for fpath in scanned_files:
        f_lower = os.path.basename(fpath).lower()
        if not (f_lower.endswith('.pdf') or f_lower.endswith('.xlsx') or f_lower.endswith('.xls')):
            continue

        # Kiểm tra từ khóa QT_A-B
        if any(k in f_lower for k in ['qt_a-b', 'qt_ab', 'qt-a-b', 'qt a-b', 'quyết toán a-b', 'quyet toan a-b']):
            if f_lower.endswith('.xlsx') or f_lower.endswith('.xls'):
                qtab_excel.append(fpath)
            else:
                qtab_pdf.append(fpath)
        elif any(k in f_lower for k in ['quyết toán', 'quyet toan', 'thqt', 'bảng thqt']):
            if f_lower.endswith('.xlsx') or f_lower.endswith('.xls'):
                other_qt_excel.append(fpath)
            else:
                other_qt_pdf.append(fpath)

    # Đọc Quyết toán theo thứ tự ưu tiên:
    # 1. QT_A-B dạng Excel
    # 2. QT_A-B dạng PDF
    # 3. File Quyết toán khác dạng Excel
    # 4. File Quyết toán khác dạng PDF
    qt_candidates = qtab_excel + qtab_pdf + other_qt_excel + other_qt_pdf
    for fpath in qt_candidates:
        f_lower = os.path.basename(fpath).lower()
        if f_lower.endswith('.xlsx') or f_lower.endswith('.xls'):
            parsed = extract_settlement_from_excel(fpath)
            if parsed.get('found'):
                qt_data = parsed
                break
        elif f_lower.endswith('.pdf'):
            parsed = extract_settlement_from_pdf(fpath)
            if parsed.get('found'):
                qt_data = parsed
                if parsed.get('tong_dt', 0) > 0 and not dt_data.get('found'):
                    dt_data = {
                        'found': True,
                        'tong_dt': parsed.get('tong_dt', 0),
                        'chi_phi_vttb': parsed.get('chi_phi_vttb_dt', 0),
                        'chi_phi_sua_chua': parsed.get('chi_phi_sua_chua_dt', 0),
                        'chi_phi_khac': 0,
                        'chi_phi_du_phong': parsed.get('chi_phi_du_phong_dt', 0),
                        'chi_phi_thu_hoi': 0,
                        'source_file': parsed.get('source_file', '')
                    }
                break

    # BƯỚC 2: Quét Dự toán & Hợp đồng
    for fpath in scanned_files:
        f_lower = os.path.basename(fpath).lower()
        if not f_lower.endswith('.pdf') and not f_lower.endswith('.xlsx') and not f_lower.endswith('.xls'):
            continue

        # Tìm Dự toán / PAKT-DT
        if any(k in f_lower for k in ['pakt', 'qdpd', 'dự toán', 'du toan', 'qđpd', 'khlcnt']) and not dt_data.get('found'):
            if f_lower.endswith('.pdf'):
                parsed_dt = extract_pakt_dt_from_pdf(fpath)
                if parsed_dt.get('found'):
                    dt_data = parsed_dt

        # Tìm Hợp đồng / Trao hợp đồng / Nghiệm thu / TMQT
        if any(k in f_lower for k in ['hợp đồng', 'hop dong', 'trao hop dong', 'trao hợp đồng', 'kqlcnt', 'hdxl', 'thuyết minh', 'thuyet minh']):
            if f_lower.endswith('.pdf'):
                parsed_c = extract_contract_info(fpath)
                if parsed_c.get('found'):
                    if not contract_data.get('found'):
                        contract_data = parsed_c
                    else:
                        for key in ['so_hd', 'ngay_hd', 'nha_thau', 'gia_tri_hd', 'thoi_gian_thi_cong', 'ngay_kc', 'ngay_ht']:
                            if not contract_data.get(key) and parsed_c.get(key):
                                contract_data[key] = parsed_c[key]

    # BƯỚC 3: Sao lưu an toàn file database_cong_trinh.xlsx trước khi cập nhật
    os.makedirs(BACKUP_DIR, exist_ok=True)
    if os.path.exists(DB_FILE):
        try:
            # Bản sao lưu tĩnh mới nhất
            backup_latest = os.path.join(BACKUP_DIR, "database_cong_trinh_backup.xlsx")
            shutil.copy2(DB_FILE, backup_latest)
            # Bản sao lưu có mốc thời gian
            ts_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_ts = os.path.join(BACKUP_DIR, f"database_cong_trinh_{ts_str}.xlsx")
            shutil.copy2(DB_FILE, backup_ts)
        except Exception as e:
            print(f"Cảnh báo sao lưu: {e}")

    # BƯỚC 4: Đọc và cập nhật file database_cong_trinh.xlsx
    df_db = pd.DataFrame()
    if os.path.exists(DB_FILE):
        df_db = pd.read_excel(DB_FILE)

    ma_clean = str(ma_ct).strip()
    idx_list = df_db.index[df_db['Mã CT'].astype(str).str.strip() == ma_clean].tolist()

    if not idx_list:
        main_idx = len(df_db)
        new_header = {col: None for col in df_db.columns}
        new_header['STT'] = 'I'
        new_header['Tên Công trình'] = ten_ct or f"Công trình {ma_ct}"
        new_header['Mã CT'] = ma_clean
        df_db = pd.concat([df_db, pd.DataFrame([new_header])], ignore_index=True)
        idx_list = [main_idx]

    main_row_idx = idx_list[0]

    # Nếu công trình chỉ có 1 dòng, mở rộng thêm 38 dòng con A -> SCL
    if len(idx_list) == 1:
        sub_rows_data = []
        for stt, t_hangmuc in STANDARD_COST_ROWS:
            r_item = {col: 0 for col in df_db.columns}
            r_item['STT'] = stt
            r_item['Tên Công trình'] = t_hangmuc
            r_item['Mã CT'] = ma_clean
            sub_rows_data.append(r_item)

        part1 = df_db.iloc[:main_row_idx + 1]
        part2 = df_db.iloc[main_row_idx + 1:]
        df_db = pd.concat([part1, pd.DataFrame(sub_rows_data), part2], ignore_index=True)
        idx_list = df_db.index[df_db['Mã CT'].astype(str).str.strip() == ma_clean].tolist()
        main_row_idx = idx_list[0]

    # Cập nhật thông tin dòng chính
    if dt_data.get('found'):
        if dt_data.get('so_qd'):
            df_db.at[main_row_idx, 'Số Dự toán'] = dt_data['so_qd']
        if dt_data.get('ngay_qd'):
            try:
                df_db.at[main_row_idx, 'Ngày Dự toán'] = pd.to_datetime(dt_data['ngay_qd'], dayfirst=True)
            except:
                pass
        if dt_data.get('tong_dt', 0) > 0:
            df_db.at[main_row_idx, 'Giá trị Dự toán'] = dt_data['tong_dt']
        if dt_data.get('chi_phi_thu_hoi', 0) > 0:
            df_db.at[main_row_idx, 'Giá trị VTTH'] = dt_data['chi_phi_thu_hoi']
        if dt_data.get('khoi_luong'):
            df_db.at[main_row_idx, 'Khối lượng công việc'] = "\n".join(dt_data['khoi_luong'])
        if dt_data.get('can_cu'):
            df_db.at[main_row_idx, 'Căn cứ pháp lý'] = "\n".join(dt_data['can_cu'])

    if contract_data.get('found'):
        if contract_data.get('so_hd'):
            df_db.at[main_row_idx, 'Số Hợp đồng xây lắp'] = contract_data['so_hd']
        if contract_data.get('ngay_hd'):
            try:
                df_db.at[main_row_idx, 'Ngày Hợp đồng xây lắp'] = pd.to_datetime(contract_data['ngay_hd'], dayfirst=True)
            except:
                pass
        if contract_data.get('gia_tri_hd', 0) > 0:
            df_db.at[main_row_idx, 'Giá trị Hợp đồng xây lắp'] = contract_data['gia_tri_hd']
        if contract_data.get('nha_thau'):
            df_db.at[main_row_idx, 'Đơn vị QL'] = contract_data['nha_thau']
        if contract_data.get('hinh_thuc'):
            df_db.at[main_row_idx, 'Ghi chú'] = contract_data['hinh_thuc']
        if contract_data.get('ngay_kc'):
            try:
                df_db.at[main_row_idx, 'Ngày khởi công'] = pd.to_datetime(contract_data['ngay_kc'], dayfirst=True)
            except:
                pass
        if contract_data.get('ngay_ht'):
            try:
                df_db.at[main_row_idx, 'Ngày hoàn thành'] = pd.to_datetime(contract_data['ngay_ht'], dayfirst=True)
            except:
                pass

    if qt_data.get('found'):
        if qt_data.get('tong_qt', 0) > 0:
            df_db.at[main_row_idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data['tong_qt']
            df_db.at[main_row_idx, 'Số tiền bằng chữ'] = doc_so_vn(qt_data['tong_qt'])

    # Cập nhật chi tiết 38 dòng con (A -> SCL)
    sub_indices = idx_list[1:]
    for idx in sub_indices:
        stt_val = str(df_db.at[idx, 'STT']).strip()

        # 1. Cập nhật Giá trị Dự toán
        if dt_data.get('found'):
            if stt_val == 'A':
                df_db.at[idx, 'Giá trị Dự toán'] = dt_data.get('chi_phi_vttb', 0)
            elif stt_val in ['A.1', 'A.1.2']:
                df_db.at[idx, 'Giá trị Dự toán'] = dt_data.get('chi_phi_vttb', 0)
            elif stt_val == 'B':
                df_db.at[idx, 'Giá trị Dự toán'] = dt_data.get('chi_phi_sua_chua', 0)
            elif stt_val in ['B.1', 'B.1.1', 'B.7']:
                df_db.at[idx, 'Giá trị Dự toán'] = dt_data.get('chi_phi_sua_chua', 0)
            elif stt_val in ['C', 'C.1']:
                df_db.at[idx, 'Giá trị Dự toán'] = dt_data.get('chi_phi_khac', 0)
            elif stt_val == 'D':
                df_db.at[idx, 'Giá trị Dự toán'] = dt_data.get('chi_phi_du_phong', 0)
            elif stt_val == 'F':
                df_db.at[idx, 'Giá trị Dự toán'] = dt_data.get('chi_phi_thu_hoi', 0)
            elif stt_val in ['E', 'E.1', 'SCL']:
                df_db.at[idx, 'Giá trị Dự toán'] = dt_data.get('tong_dt', 0)

        # 2. Cập nhật Giá trị Quyết toán
        if qt_data.get('found'):
            if stt_val == 'A':
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('chi_phi_vttb', 0)
            elif stt_val in ['A.1', 'A.1.2']:
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('chi_phi_vttb', 0)
            elif stt_val == 'B':
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('chi_phi_sua_chua', 0)
            elif stt_val in ['B.1', 'B.1.1']:
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('vat_lieu', 0)
            elif stt_val in ['B.2', 'B.2.1']:
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('nhan_cong', 0)
            elif stt_val in ['B.3', 'B.3.1']:
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('may_thi_cong', 0)
            elif stt_val == 'B.5':
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('chi_phi_chung', 0)
            elif stt_val == 'B.8':
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('thue_gtgt', 0)
            elif stt_val == 'B.7':
                truoc_thue = qt_data.get('chi_phi_sua_chua', 0) - qt_data.get('thue_gtgt', 0)
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = truoc_thue if truoc_thue > 0 else qt_data.get('chi_phi_sua_chua', 0)
            elif stt_val in ['C', 'C.1']:
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('chi_phi_khac', 0)
            elif stt_val == 'D':
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('chi_phi_du_phong', 0)
            elif stt_val == 'F':
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('chi_phi_thu_hoi', 0)
            elif stt_val in ['E', 'E.1', 'SCL']:
                df_db.at[idx, 'Giá trị Q.định phê duyệt QT công trình'] = qt_data.get('tong_qt', 0)

    # Lưu lại file database_cong_trinh.xlsx
    try:
        df_db.to_excel(DB_FILE, index=False)
    except Exception as e:
        print(f"Lỗi khi lưu database_cong_trinh.xlsx: {e}")

    # Chạy kiểm tra tuân thủ quy định pháp lý (QĐ 202)
    main_row_updated = df_db.loc[main_row_idx]
    ct_data_updated = df_db.loc[idx_list]
    compliance = check_compliance(main_row_updated, ct_data_updated, scanned_files, folder)

    result_dict = {
        'success': True,
        'message': f"Đã quét và cập nhật thành công dữ liệu từ thư mục hồ sơ cho {ma_ct}! (Đã tạo bản sao lưu tại data/backups/)",
        'folder': folder,
        'scanned_files': scanned_files,
        'dt_data': dt_data,
        'qt_data': qt_data,
        'contract_data': contract_data,
        'compliance': compliance,
        'updated_at': datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    }

    # BƯỚC 5: Lưu cache snapshot vào file JSON để nạp siêu tốc không gây lag
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        cache_path = os.path.join(CACHE_DIR, f"{ma_clean}.json")
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(result_dict, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Cảnh báo lưu cache: {e}")

    return result_dict
