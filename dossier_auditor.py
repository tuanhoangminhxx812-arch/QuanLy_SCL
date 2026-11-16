"""
dossier_auditor.py
Module Chuyên sâu Quản lý, Kiểm tra Hồ sơ & Đối chiếu Vật tư SCL theo Quyết định 202/QĐ-HĐTV (EVNHCMC)
Cung cấp:
1. Quét & Phân loại 4 nhóm hồ sơ QĐ 202 (nhận diện PDF text, cảnh báo PDF scan chưa OCR, đọc Excel)
2. Trích xuất thông tin pháp lý, số QĐ, ngày tháng & rà soát tính hợp lệ của trình tự thời gian (Timeline)
3. Module Đối chiếu chéo Bảng XNKL (khối lượng nghiệm thu) và Bảng Quyết toán Vật tư A cấp
4. Xuất Báo cáo Kiểm tra SCL chuẩn EVN ra file Excel BaoCao_KiemTra_SCL_QD202.xlsx (4 sheet, Times New Roman 12)
"""

import os
import re
import datetime
from io import BytesIO
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HO_SO_DIR = os.path.join(BASE_DIR, "Ho_so_cong_trinh")

# ── 1. ĐỊNH NGHĨA 4 NHÓM HỒ SƠ CHUẨN QĐ 202/QĐ-HĐTV ──
QD202_GROUPS = {
    1: {
        "code": "NHOM_1",
        "name": "1. Pháp lý & Dự toán",
        "desc": "Kế hoạch vốn, danh mục SCL, khảo sát hiện trạng, phê duyệt PAKT-Dự toán",
        "required_docs": [
            ("QD_KH_VON", "Quyết định giao KH vốn / Văn bản chấp thuận danh mục SCL", ["chấp thuận", "ke hoach", "giao ke hoach", "danh muc", "cap von", "2635"]),
            ("BB_KHAO_SAT", "Biên bản khảo sát hư hỏng / đánh giá hiện trạng", ["khảo sát", "khao sat", "hư hỏng", "hu hong", "cbm", "rcm", "hien trang"]),
            ("QD_PAKT_DT", "Quyết định phê duyệt PAKT - Dự toán", ["qdpd pakt", "phe duyet pakt", "qd pakt", "pakt-dt", "pakt", "tong du toan", "dự toán", "du toan"]),
        ]
    },
    2: {
        "code": "NHOM_2",
        "name": "2. Thực hiện, Hợp đồng & Vật tư",
        "desc": "KHLCNT / Thỏa thuận giao việc, Hợp đồng xây lắp & mua sắm, cấp phát VTTB",
        "required_docs": [
            ("KHLCNT_TTGV", "Kế hoạch LCNT / Thỏa thuận giao việc (nếu tự làm)", ["khlcnt", "ke hoach lua chon", "thoa thuan giao viec", "ttgv", "giao nhiem vu"]),
            ("HOP_DONG_XL", "Hợp đồng xây lắp / thi công (hoặc TTGV nếu tự làm)", ["hdxl", "hợp đồng xây lắp", "hop dong xay lap", "hd-", "hđ-", "hđxl", "hđptv"]),
            ("HOP_DONG_MS", "Hợp đồng mua sắm VTTB / Thư chấp thuận vật tư", ["mua sắm", "mua sam", "vật tư", "vat tu", "giao nhận hàng hóa", "giao nhan hang hoa", "thư chấp thuận"]),
            ("PHIEU_XUAT_KHO", "Biên bản giao nhận / Phiếu xuất kho VTTB A cấp", ["giao nhận", "giao nhan", "xuất kho", "xuat kho", "phieu xuat", "bbgn"]),
        ]
    },
    3: {
        "code": "NHOM_3",
        "name": "3. Kỹ thuật & Nghiệm thu",
        "desc": "Thông báo khởi công, giám sát, biên bản nghiệm thu kỹ thuật, bảng XNKL",
        "required_docs": [
            ("TB_KHOI_CONG", "Thông báo khởi công công trình", ["khoi cong", "khởi công", "thông báo khởi công"]),
            ("QD_GIAM_SAT", "Quyết định / Biên bản cử giám sát thi công", ["giam sat", "giám sát", "cử giám sát", "nhiem vu giam sat"]),
            ("BBNT_KY_THUAT", "Biên bản nghiệm thu kỹ thuật / thử nghiệm", ["nt ky thuat", "nghiệm thu kỹ thuật", "nghiem thu ky thuat", "thử nghiệm", "thu nghiem", "suat tieu hao", "vttb truoc khi"]),
            ("BANG_XNKL", "Bảng xác nhận khối lượng hoàn thành (Bảng XNKL)", ["xnkl", "xác nhận khối lượng", "xac nhan khoi luong", "kl hoàn thành", "kl hoan thanh", "bbnt kl"]),
            ("BBNT_HOAN_THANH", "Biên bản nghiệm thu đưa vào sử dụng / hoàn thành", ["dua vao su dung", "đưa vào sử dụng", "hoàn thành", "hoan thanh", "bbnt hoan thanh"]),
        ]
    },
    4: {
        "code": "NHOM_4",
        "name": "4. Quyết toán & Thu hồi",
        "desc": "Quyết toán A-B, quyết toán vật tư A cấp, thu hồi VTTB thừa/hỏng, thẩm tra QT",
        "required_docs": [
            ("QT_A_B", "Bảng quyết toán A-B / Bảng THQT KP SCL", ["qt_a-b", "quyết toán a-b", "quyet toan a-b", "bảng thqt", "bang thqt", "ho so quyet toan chot", "bia qt", "qtth"]),
            ("QT_VTTB_A_CAP", "Bảng quyết toán vật tư, thiết bị A cấp cho ĐVTC", ["qt vttb", "quyết toán vật tư", "quyet toan vat tu", "vttb a cấp", "qt cp tb vt"]),
            ("BB_THU_HOI_VT", "Biên bản đánh giá / Bảng kê VTTB thu hồi nhập kho", ["thu hồi", "thu hoi", "vtth", "danh gia nhap kho", "danh gia vttb thu hoi", "bbdg vtth"]),
            ("HOA_DON_GTGT", "Hóa đơn GTGT (nếu thuê ngoài thi công / mua sắm)", ["hóa đơn", "hoa don", "invoice", "vat", "039", "189"]),
            ("TO_TRINH_DUYET_QT", "Tờ trình duyệt quyết toán danh mục SCL", ["tờ trình", "to trinh", "ttr", "đề nghị duyệt"]),
            ("BC_THAM_TRA", "Báo cáo kết quả thẩm tra quyết toán SCL", ["báo cáo thẩm tra", "bao cao tham tra", "bc tham tra", "tham tra qt"]),
            ("QD_PHE_DUYET_QT", "Quyết định phê duyệt quyết toán SCL", ["qđ pdqt", "phê duyệt quyết toán", "phe duyet quyet toan", "qd phe duyet"]),
        ]
    }
}


# ── 2. TỪ ĐIỂN CHUẨN HÓA VẬT TƯ (MATERIAL NORMALIZATION DICTIONARY) ──
MATERIAL_ALIASES = [
    (r"\b(?:cầu\s*chì\s*tự\s*rơi|fco)\b(?:\s*-\s*\d+a|\s*22kv|\s*24kv)?", "FCO 24kV"),
    (r"\b(?:cầu\s*chì\s*cắt\s*có\s*tải|lbfco)\b(?:\s*-\s*\d+a|\s*22kv|\s*24kv)?", "LBFCO 24kV"),
    (r"\b(?:chống\s*sét\s*van|chong\s*set\s*van|la)\b(?:\s*18kv|\s*18|\s*24kv|\s*24)?", "Chống sét van LA 18kV"),
    (r"\b(?:chụp\s*đầu\s*cực|nắp\s*chụp|chup\s*dau\s*cuc|chụp\s*silicon)\b[^\n\r,;]*", "Chụp đầu cực FCO/LBFCO/LA"),
    (r"\b(?:dây\s*chì|fuse\s*link|day\s*chi)\b\s*(\d+[kK])", r"Dây chì Fuse link \1"),
    (r"\b(?:cáp\s*đồng\s*bọc|cap\s*dong\s*boc|cu/pvc\s*25)\b\s*(?:25\s*mm2|25mm2|25)?", "Cáp đồng bọc 25mm2"),
    (r"\b(?:kẹp\s*hotline|kep\s*hotline)\b\s*(?:25-70|25/70)?", "Kẹp hotline 25-70"),
    (r"\b(?:kẹp\s*quai|kep\s*quai)\b", "Kẹp quai"),
    # Phụ tùng máy phát điện Cummins G1, G2
    (r"\b(?:bộ\s*kit\s*xi\s*-?\s*lanh|xi\s*-?\s*lanh)\b", "Bộ kit Xi-lanh"),
    (r"\b(?:bạc\s*xéc\s*-?\s*măng|xéc\s*-?\s*măng)\b", "Bạc Xéc-măng"),
    (r"\b(?:bộ\s*bạc\s*trục\s*khuỷu|trục\s*khuỷu)\b", "Bộ Bạc trục khuỷu"),
    (r"\b(?:bạc\s*đầu\s*to\s*thanh\s*truyền)\b", "Bạc đầu to thanh truyền"),
    (r"\b(?:bộ\s*ron\s*các\s*-?\s*te|ron\s*các\s*-?\s*te)\b", "Bộ ron các-te"),
    (r"\b(?:bộ\s*sửa\s*chữa\s*turbo|sửa\s*chữa\s*turbo)\b", "Bộ sửa chữa turbo"),
    (r"\b(?:ron\s*turbo)\b", "Ron turbo"),
    (r"\b(?:ron\s*ống\s*nhớt\s*hồi\s*turbo)\b", "Ron ống nhớt hồi turbo"),
    (r"\b(?:ron\s*cụm\s*khí\s*nạp)\b", "Ron cụm khí nạp"),
    (r"\b(?:ron\s*lõi\s*sinh\s*hàn\s*gió)\b", "Ron lõi sinh hàn gió"),
    (r"\b(?:ron\s*đường\s*nước\s*lõi)\b", "Ron đường nước lõi"),
]


def normalize_material_name(raw_name):
    """Chuẩn hóa tên vật tư về danh mục chuẩn thống nhất."""
    if not raw_name:
        return ""
    name = str(raw_name).strip()
    name_clean = re.sub(r'\s+', ' ', name).strip()
    for pattern, std_name in MATERIAL_ALIASES:
        if re.search(pattern, name_clean, re.IGNORECASE):
            # Check backreference substitution
            try:
                return re.sub(pattern, std_name, name_clean, flags=re.IGNORECASE).strip()
            except:
                return std_name
    return name_clean


def inspect_pdf_text_and_ocr(file_path):
    """
    Kiểm tra file PDF: trích xuất text, số trang, và kiểm tra xem có phải scan ảnh chưa OCR không.
    Trả về: (full_text, pages_count, is_scanned)
    """
    full_text = ""
    pages_count = 0
    is_scanned = False

    # Thử pdfplumber trước
    if pdfplumber is not None:
        try:
            with pdfplumber.open(file_path) as pdf:
                pages_count = len(pdf.pages)
                for p in pdf.pages:
                    t = p.extract_text() or ""
                    full_text += t + "\n"
        except Exception:
            pass

    # Fallback sang pypdf nếu pdfplumber chưa đọc được
    if not full_text and PdfReader is not None:
        try:
            reader = PdfReader(file_path)
            pages_count = len(reader.pages)
            for p in reader.pages:
                t = p.extract_text() or ""
                full_text += t + "\n"
        except Exception:
            pass

    # Đánh giá xem có phải scan ảnh chưa OCR không
    clean_len = len(full_text.strip())
    if pages_count > 0 and clean_len < 35 * pages_count:
        is_scanned = True

    return full_text, pages_count, is_scanned


def extract_metadata_from_text(text, filename=""):
    """
    Trích xuất Mã CT (VTAD...), Số QĐ / HĐ / TTGV, và ngày tháng từ văn bản hoặc tên file.
    """
    combined = f"{filename}\n{text}"
    ma_ct = ""
    so_qd = ""
    ngay_van_ban = None

    # Mã công trình VTAD...
    m_ma = re.search(r'VTAD\d{7}', combined, re.IGNORECASE)
    if m_ma:
        ma_ct = m_ma.group(0).upper()

    # Số Quyết định / Số HĐ / Số TTGV
    m_so = re.search(r'(?:Số|Số:)\s*([^\n\r/]+(?:/[^\n\r,]+)?)', text, re.IGNORECASE)
    if m_so:
        s_val = m_so.group(1).strip()
        if len(s_val) >= 3 and not s_val.startswith('/'):
            so_qd = s_val.split()[0] if ' ' in s_val and '/' not in s_val else s_val
    if not so_qd:
        # Tìm từ tên file ví dụ "2132-thư chấp thuận", "352-QD- Giao..."
        m_fn = re.search(r'^(\d+[\-_A-Za-z0-9/]+)', filename)
        if m_fn:
            so_qd = m_fn.group(1).strip('-_ ')

    # Ngày tháng năm
    m_date = re.search(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})', text, re.IGNORECASE)
    if m_date:
        d, m, y = m_date.groups()
        try:
            ngay_van_ban = datetime.date(int(y), int(m), int(d))
        except:
            pass
    if not ngay_van_ban:
        m_d2 = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', combined)
        if m_d2:
            d, m, y = m_d2.groups()
            try:
                ngay_van_ban = datetime.date(int(y), int(m), int(d))
            except:
                pass

    return {
        "ma_ct": ma_ct,
        "so_qd": so_qd,
        "ngay": ngay_van_ban,
        "ngay_str": ngay_van_ban.strftime("%d/%m/%Y") if ngay_van_ban else ""
    }


def classify_file_into_qd202(filename, text="", is_scanned=False):
    """
    Phân loại một file vào 1 trong 4 nhóm hồ sơ QĐ 202 và chỉ ra doc_type khớp nhất.
    Ưu tiên 1: Tên file (Filename) có độ chính xác cao nhất.
    Ưu tiên 2: Nội dung văn bản (Text) với cụm từ đặc trưng rõ ràng.
    """
    fn_lower = filename.lower()
    txt_lower = text.lower() if text else ""

    # ── BƯỚC 1: SO KHỚP ƯU TIÊN THEO TÊN FILE ──
    # Nhóm 3: Khởi công & Giám sát (rất đặc trưng)
    if any(k in fn_lower for k in ["khoi cong", "khởi công"]):
        return 3, "TB_KHOI_CONG", "Thông báo khởi công công trình"
    if any(k in fn_lower for k in ["giam sat", "giám sát"]):
        return 3, "QD_GIAM_SAT", "Quyết định / Biên bản cử giám sát thi công"

    # Nhóm 4: Hóa đơn & Thu hồi & Quyết toán
    if any(k in fn_lower for k in ["hoa don", "hóa đơn", "invoice", "1c26", "c26", "vat", "3500822067"]):
        return 4, "HOA_DON_GTGT", "Hóa đơn GTGT (nếu thuê ngoài thi công / mua sắm)"
    if any(k in fn_lower for k in ["thu hoi", "thu hồi", "vtth"]):
        return 4, "BB_THU_HOI_VT", "Biên bản đánh giá / Bảng kê VTTB thu hồi nhập kho"
    if any(k in fn_lower for k in ["bang th qt", "th qt", "thqt", "tong hop scl", "quyet toan", "quyết toán", "qtda", "bo quyet toan", "ho so quyet toan", "de nghi thanh toan"]):
        return 4, "QT_A_B", "Bảng quyết toán A-B / Bảng THQT KP SCL"
    if any(k in fn_lower for k in ["chiet tinh dv6", "chiết tính dv6", "82 vt", "qt vttb", "vttb a cấp"]):
        return 4, "QT_VTTB_A_CAP", "Bảng quyết toán vật tư, thiết bị A cấp cho ĐVTC"
    if any(k in fn_lower for k in ["tham tra", "thẩm tra"]):
        return 4, "BC_THAM_TRA", "Báo cáo kết quả thẩm tra quyết toán SCL"
    if any(k in fn_lower for k in ["to trinh", "tờ trình", "ttr"]):
        return 4, "TO_TRINH_DUYET_QT", "Tờ trình duyệt quyết toán danh mục SCL"

    # Nhóm 1: Khảo sát, Kế hoạch vốn, Phê duyệt PAKT - Dự toán
    if any(k in fn_lower for k in ["khao sat", "khảo sát", "hu hong", "hư hỏng", "cbm", "rcm", "kstthh", "ks tthh", "khiem khuyet", "khiếm khuyết"]):
        return 1, "BB_KHAO_SAT", "Biên bản khảo sát hư hỏng / đánh giá hiện trạng"
    if any(k in fn_lower for k in ["chap thuan", "chấp thuận", "ke hoach von", "kế hoạch vốn", "giao ke hoach", "danh muc scl", "2132"]):
        return 1, "QD_KH_VON", "Quyết định giao KH vốn / Văn bản chấp thuận danh mục SCL"
    if any(k in fn_lower for k in ["pa kt dt", "pa kt", "pa_kt", "pakt", "dự toán", "du toan"]):
        return 1, "QD_PAKT_DT", "Quyết định phê duyệt PAKT - Dự toán"

    # Nhóm 3: Nghiệm thu kỹ thuật & Đưa vào sử dụng & Bảng XNKL
    if any(k in fn_lower for k in ["chiet tinh tung cong viec", "xnkl", "xac nhan khoi luong", "xác nhận khối lượng", "ntkl"]):
        return 3, "BANG_XNKL", "Bảng xác nhận khối lượng hoàn thành (Bảng XNKL)"
    if any(k in fn_lower for k in ["dua vao su dung", "đưa vào sử dụng", "bbnt va", "bbnt", "hoan thanh", "hoàn thành"]):
        return 3, "BBNT_HOAN_THANH", "Biên bản nghiệm thu đưa vào sử dụng / hoàn thành"
    if any(k in fn_lower for k in ["nt ky thuat", "nghiem thu ky thuat", "nghiệm thu kỹ thuật", "suat tieu hao", "truoc khi su dung", "ky thuat"]):
        return 3, "BBNT_KY_THUAT", "Biên bản nghiệm thu kỹ thuật / thử nghiệm"

    # Nhóm 2: Hợp đồng, Thỏa thuận, Phiếu xuất kho
    if any(k in fn_lower for k in ["ttgv", "thoa thuan", "thỏa thuận", "phu_luc_ttgv", "phu luc ttgv", "giao việc", "giao viec"]):
        return 2, "KHLCNT_TTGV", "Thỏa thuận giao việc (TTGV) / Kế hoạch LCNT"
    if any(k in fn_lower for k in ["hop dong", "hợp đồng", "hd-", "hđ-", "hđxl", "hdxl"]):
        if any(k in fn_lower for k in ["mua sam", "mua sắm", "vật tư", "vat tu", "shin", "xpad"]):
            return 2, "HOP_DONG_MS", "Hợp đồng mua sắm VTTB / Thư chấp thuận vật tư"
        return 2, "HOP_DONG_XL", "Hợp đồng xây lắp / thi công (hoặc TTGV nếu tự làm)"
    if any(k in fn_lower for k in ["giao nhan", "giao nhận", "xuat kho", "xuất kho", "bbgn", "phieu xuat"]):
        return 2, "PHIEU_XUAT_KHO", "Biên bản giao nhận / Phiếu xuất kho VTTB A cấp"

    # ── BƯỚC 2: NẾU TÊN FILE CHƯA KHỚP, XÉT NỘI DUNG VĂN BẢN (TEXT) ──
    if txt_lower and not is_scanned:
        if any(k in txt_lower for k in ["hóa đơn giá trị gia tăng", "hoa don gia tri gia tang", "mẫu số 01gtkt", "mẫu số 1c26"]):
            return 4, "HOA_DON_GTGT", "Hóa đơn GTGT (nếu thuê ngoài thi công / mua sắm)"
        if any(k in txt_lower for k in ["thông báo khởi công", "thong bao khoi cong"]):
            return 3, "TB_KHOI_CONG", "Thông báo khởi công công trình"
        if any(k in txt_lower for k in ["giao nhiệm vụ giám sát", "cử giám sát thi công", "giám sát thi công xây dựng"]):
            return 3, "QD_GIAM_SAT", "Quyết định / Biên bản cử giám sát thi công"
        if any(k in txt_lower for k in ["đưa vào sử dụng", "nghiệm thu hoàn thành hạng mục"]):
            return 3, "BBNT_HOAN_THANH", "Biên bản nghiệm thu đưa vào sử dụng / hoàn thành"
        if any(k in txt_lower for k in ["nghiệm thu kỹ thuật", "suất tiêu hao nhiên liệu", "vttb trước khi sử dụng"]):
            return 3, "BBNT_KY_THUAT", "Biên bản nghiệm thu kỹ thuật / thử nghiệm"
        if any(k in txt_lower for k in ["biên bản giao nhận hàng hóa", "phiếu xuất kho"]):
            return 2, "PHIEU_XUAT_KHO", "Biên bản giao nhận / Phiếu xuất kho VTTB A cấp"
        if any(k in txt_lower for k in ["đánh giá vật tư thu hồi", "bàn giao vttb thu hồi"]):
            return 4, "BB_THU_HOI_VT", "Biên bản đánh giá / Bảng kê VTTB thu hồi nhập kho"
        if any(k in txt_lower for k in ["bảng tổng hợp quyết toán", "quyết toán công trình hoàn thành"]):
            return 4, "QT_A_B", "Bảng quyết toán A-B / Bảng THQT KP SCL"
        if any(k in txt_lower for k in ["phê duyệt phương án kỹ thuật", "phê duyệt pakt", "phê duyệt dự toán"]):
            return 1, "QD_PAKT_DT", "Quyết định phê duyệt PAKT - Dự toán"
        if any(k in txt_lower for k in ["chấp thuận danh mục", "giao kế hoạch vốn"]):
            return 1, "QD_KH_VON", "Quyết định giao KH vốn / Văn bản chấp thuận danh mục SCL"
        if any(k in txt_lower for k in ["hợp đồng thi công", "hợp đồng kinh tế", "hợp đồng số"]):
            return 2, "HOP_DONG_XL", "Hợp đồng xây lắp / thi công (hoặc TTGV nếu tự làm)"

    return None, "CHUA_PHAN_LOAI", "Tài liệu khác / Chưa phân loại"


# ── 3. QUÉT TOÀN DIỆN THƯ MỤC CÔNG TRÌNH THEO QĐ 202 ──
def scan_project_dossier_qd202(project_folder, ma_ct="", ten_ct="", is_tu_lam=False):
    """
    Quét đệ quy toàn bộ thư mục công trình, phân loại 4 nhóm, trích xuất timeline,
    gắn cờ cảnh báo OCR và cờ thiếu sót hồ sơ.
    """
    scanned_files = []
    if not project_folder or not os.path.exists(project_folder):
        return {
            "folder": project_folder,
            "total_files": 0,
            "groups": {1: [], 2: [], 3: [], 4: []},
            "unclassified": [],
            "scanned_pdf_count": 0,
            "timeline": [],
            "missing_docs": [],
            "warnings": ["Chưa tìm thấy thư mục hồ sơ công trình trên máy."],
            "is_tu_lam": is_tu_lam
        }

    for root, _, files in os.walk(project_folder):
        for f in files:
            scanned_files.append(os.path.join(root, f))

    groups_data = {1: [], 2: [], 3: [], 4: []}
    unclassified_files = []
    timeline_events = []
    scanned_pdf_count = 0
    all_files_info = []

    for fpath in scanned_files:
        fname = os.path.basename(fpath)
        ext = os.path.splitext(fname)[1].lower()
        if ext in [".tmp", ".lnk", ".bak"]:
            continue

        text = ""
        pages = 1
        is_scanned = False

        sheet_names = []
        if ext in [".xlsx", ".xls", ".xlsb"]:
            # Đọc sơ bộ danh sách sheet
            try:
                if ext == ".xlsb":
                    xl = pd.ExcelFile(fpath, engine="pyxlsb")
                else:
                    xl = pd.ExcelFile(fpath)
                sheet_names = xl.sheet_names
                text = " ".join(sheet_names)
            except:
                text = ""

        meta = extract_metadata_from_text(text, fname)
        gid, doc_key, doc_name = classify_file_into_qd202(fname, text, is_scanned)

        rel_path = os.path.relpath(fpath, project_folder)
        file_obj = {
            "name": fname,
            "path": fpath,
            "rel_path": rel_path,
            "ext": ext,
            "pages": pages,
            "is_scanned": is_scanned,
            "group_id": gid,
            "doc_key": doc_key,
            "doc_name": doc_name,
            "ma_ct": meta["ma_ct"],
            "so_qd": meta["so_qd"],
            "ngay": meta["ngay"],
            "ngay_str": meta["ngay_str"],
            "sheet_names": sheet_names,
        }
        all_files_info.append(file_obj)

        if gid in groups_data:
            groups_data[gid].append(file_obj)
        else:
            unclassified_files.append(file_obj)

        if meta["ngay"]:
            timeline_events.append({
                "date": meta["ngay"],
                "date_str": meta["ngay_str"],
                "group_id": gid or 0,
                "doc_name": doc_name,
                "so_qd": meta["so_qd"],
                "filename": fname
            })

    # Sắp xếp timeline theo thứ tự thời gian
    timeline_events.sort(key=lambda x: x["date"])

    # Kiểm tra tính logic thời gian & tìm cảnh báo
    timeline_warnings = []
    # Ví dụ: kiểm tra ngày khởi công vs ngày hợp đồng
    date_hd = None
    date_kc = None
    date_nt = None
    date_qt = None

    for ev in timeline_events:
        d_name_l = ev["doc_name"].lower()
        if "hợp đồng" in d_name_l or "thỏa thuận giao việc" in d_name_l:
            if not date_hd: date_hd = ev["date"]
        elif "khởi công" in d_name_l:
            if not date_kc: date_kc = ev["date"]
        elif "nghiệm thu" in d_name_l:
            date_nt = ev["date"]
        elif "quyết toán" in d_name_l:
            date_qt = ev["date"]

    if date_hd and date_kc and date_kc < date_hd:
        timeline_warnings.append(f"⚠️ Rủi ro trình tự: Ngày khởi công ({date_kc.strftime('%d/%m/%Y')}) trước Ngày ký hợp đồng / TTGV ({date_hd.strftime('%d/%m/%Y')}).")
    if date_nt and date_qt and date_qt < date_nt:
        timeline_warnings.append(f"⚠️ Rủi ro trình tự: Ngày lập quyết toán ({date_qt.strftime('%d/%m/%Y')}) trước Ngày nghiệm thu ({date_nt.strftime('%d/%m/%Y')}).")

    # Kiểm tra hồ sơ còn thiếu theo 4 nhóm
    missing_docs = []
    present_keys = set([f["doc_key"] for f in all_files_info if f.get("doc_key") != "CHUA_PHAN_LOAI"])

    # Nếu có hợp đồng liên danh / mua sắm & thi công trọn gói hoặc TTGV
    for f in all_files_info:
        fn_l = f.get("name", "").lower()
        if any(k in fn_l for k in ["hd-", "hđ-", "hợp đồng", "hop dong"]):
            if any(k in fn_l for k in ["ld", "liên danh", "shin-xpad", "thi công", "xây lắp"]):
                present_keys.add("HOP_DONG_XL")
                present_keys.add("HOP_DONG_MS")
        if any(k in fn_l for k in ["ttgv", "thoa thuan", "thỏa thuận"]):
            present_keys.add("KHLCNT_TTGV")
            present_keys.add("HOP_DONG_XL")
        if "bbnt" in fn_l and any(k in fn_l for k in ["th qt", "thqt", "quyết toán", "quyet toan"]):
            present_keys.add("BBNT_HOAN_THANH")
            present_keys.add("QT_A_B")
            groups_data[4].append({
                "name": f["name"],
                "path": f["path"],
                "rel_path": f["rel_path"],
                "ext": f["ext"],
                "pages": 1,
                "is_scanned": f.get("is_scanned", False),
                "group_id": 4,
                "doc_key": "QT_A_B",
                "doc_name": "Bảng quyết toán A-B / Bảng THQT KP SCL",
                "so_qd": f["so_qd"],
                "ngay": f["ngay"],
                "ngay_str": f["ngay_str"]
            })
        if any(k in fn_l for k in ["chiet tinh", "chiết tính"]):
            present_keys.add("QT_VTTB_A_CAP")
            present_keys.add("BANG_XNKL")

    # Bổ sung các chứng từ được nhúng bên trong các sheet của file Excel tổng hợp
    for f in all_files_info:
        if f.get("sheet_names"):
            s_lowers = [s.lower() for s in f["sheet_names"]]
            fname = f["name"]
            for s, s_orig in zip(s_lowers, f["sheet_names"]):
                if any(k in s for k in ['bbnt kl', 'xnkl', 'xac nhan khoi luong', 'ntkl', 'bbxngt']):
                    present_keys.add('BANG_XNKL')
                    groups_data[3].append({
                        "name": f"{fname} [Sheet: '{s_orig}']",
                        "path": f["path"],
                        "rel_path": f["rel_path"],
                        "ext": f["ext"],
                        "pages": 1,
                        "is_scanned": False,
                        "group_id": 3,
                        "doc_key": "BANG_XNKL",
                        "doc_name": "Bảng xác nhận khối lượng hoàn thành (Bảng XNKL)",
                        "so_qd": f["so_qd"],
                        "ngay": f["ngay"],
                        "ngay_str": f["ngay_str"]
                    })
                if any(k in s for k in ['bbgn', 'giao nhan', 'xuất kho', 'giao nhận']):
                    present_keys.add('PHIEU_XUAT_KHO')
                    groups_data[2].append({
                        "name": f"{fname} [Sheet: '{s_orig}']",
                        "path": f["path"],
                        "rel_path": f["rel_path"],
                        "ext": f["ext"],
                        "pages": 1,
                        "is_scanned": False,
                        "group_id": 2,
                        "doc_key": "PHIEU_XUAT_KHO",
                        "doc_name": "Biên bản giao nhận / Phiếu xuất kho VTTB A cấp",
                        "so_qd": f["so_qd"],
                        "ngay": f["ngay"],
                        "ngay_str": f["ngay_str"]
                    })
                if any(k in s for k in ['vtth', 'thu hoi', 'thu hồi', 'bbdg']):
                    present_keys.add('BB_THU_HOI_VT')
                    groups_data[4].append({
                        "name": f"{fname} [Sheet: '{s_orig}']",
                        "path": f["path"],
                        "rel_path": f["rel_path"],
                        "ext": f["ext"],
                        "pages": 1,
                        "is_scanned": False,
                        "group_id": 4,
                        "doc_key": "BB_THU_HOI_VT",
                        "doc_name": "Biên bản đánh giá / Bảng kê VTTB thu hồi nhập kho",
                        "so_qd": f["so_qd"],
                        "ngay": f["ngay"],
                        "ngay_str": f["ngay_str"]
                    })
                if any(k in s for k in ['xpad', 'shin', 'vttb', 'phat sinh tang', 'tách vttb']):
                    present_keys.add('QT_VTTB_A_CAP')
                    groups_data[4].append({
                        "name": f"{fname} [Sheet: '{s_orig}']",
                        "path": f["path"],
                        "rel_path": f["rel_path"],
                        "ext": f["ext"],
                        "pages": 1,
                        "is_scanned": False,
                        "group_id": 4,
                        "doc_key": "QT_VTTB_A_CAP",
                        "doc_name": "Bảng quyết toán vật tư, thiết bị A cấp cho ĐVTC",
                        "so_qd": f["so_qd"],
                        "ngay": f["ngay"],
                        "ngay_str": f["ngay_str"]
                    })
                if any(k in s for k in ['quyết toán', 'quyet toan', 'thqt', 'qtth', 'bia qt', 'gtqt', '3a nđ254']):
                    present_keys.add('QT_A_B')

    for g_id, g_info in QD202_GROUPS.items():
        for doc_key, doc_name, _ in g_info["required_docs"]:
            # Xử lý trường hợp Tự làm
            if is_tu_lam and doc_key == "HOP_DONG_XL":
                # Tự làm không cần HĐ xây lắp nhưng bắt buộc phải có Thỏa thuận giao việc (TTGV)
                if "KHLCNT_TTGV" not in present_keys:
                    missing_docs.append({
                        "group_id": g_id,
                        "group_name": g_info["name"],
                        "doc_key": "TTGV_BAT_BUOC",
                        "doc_name": "Thỏa thuận giao việc (TTGV) - Bắt buộc đối với công trình Tự làm!",
                        "severity": "DANGER"
                    })
                continue
            if is_tu_lam and doc_key == "KHLCNT_TTGV":
                continue

            if doc_key not in present_keys:
                # Mức độ cảnh báo
                sev = "DANGER" if doc_key in ["QD_PAKT_DT", "HOP_DONG_XL", "BANG_XNKL", "QT_A_B", "BB_THU_HOI_VT"] else "WARNING"
                missing_docs.append({
                    "group_id": g_id,
                    "group_name": g_info["name"],
                    "doc_key": doc_key,
                    "doc_name": doc_name,
                    "severity": sev
                })

    return {
        "folder": project_folder,
        "total_files": len(all_files_info),
        "all_files": all_files_info,
        "groups": groups_data,
        "unclassified": unclassified_files,
        "scanned_pdf_count": scanned_pdf_count,
        "timeline": timeline_events,
        "timeline_warnings": timeline_warnings,
        "missing_docs": missing_docs,
        "is_tu_lam": is_tu_lam
    }


# ── 4. MODULE ĐỐI CHIẾU CHÉO VẬT TƯ A CẤP VÀ BẢNG XNKL ──
def reconcile_materials_scl(project_folder):
    """
    Tự động nhận diện bảng biểu từ Excel / PDF:
    - Bảng XNKL (khối lượng nghiệm thu hoàn thành)
    - Bảng Quyết toán Vật tư A cấp / Biên bản giao nhận hàng hóa
    - Bảng đánh giá Vật tư thu hồi nhập kho
    Chuẩn hóa tên vật tư và lập bảng so sánh đối chiếu chéo.
    """
    res = {
        "found": False,
        "source_xnkl": "",
        "source_vttb": "",
        "source_vtth": "",
        "items": [],
        "matched_count": 0,
        "surplus_count": 0,
        "deficit_count": 0,
        "warnings": []
    }

    if not project_folder or not os.path.exists(project_folder):
        return res

    # 1. Tìm file Excel quyết toán trước (ưu tiên số 1)
    excel_files = []
    pdf_files = []
    for root, _, files in os.walk(project_folder):
        for f in files:
            fp = os.path.join(root, f)
            if f.lower().endswith((".xlsx", ".xls")):
                excel_files.append(fp)
            elif f.lower().endswith(".pdf"):
                pdf_files.append(fp)

    xnkl_dict = {}     # {std_name: {"raw_name": ..., "dvt": ..., "sl_nt": ...}}
    vttb_dict = {}     # {std_name: {"raw_name": ..., "dvt": ..., "sl_xuat": ..., "sl_qt": ...}}
    vtth_dict = {}     # {std_name: {"raw_name": ..., "dvt": ..., "sl_thu_hoi": ...}}

    # Sắp xếp file Excel: ưu tiên file có ngày tháng mới nhất trong tên hoặc mtime mới nhất
    excel_files.sort(key=lambda x: (os.path.basename(x).startswith("31"), os.path.getmtime(x)), reverse=True)

    # Chọn 1 file Excel quyết toán chính để tránh cộng dồn số liệu giữa các file nháp cũ và mới
    primary_excel = None
    for ef in excel_files:
        try:
            xl = pd.ExcelFile(ef)
            s_lowers = [s.lower() for s in xl.sheet_names]
            if any("kl" in s or "xnkl" in s or "bbnt" in s for s in s_lowers):
                primary_excel = ef
                break
        except:
            pass
    if not primary_excel and excel_files:
        primary_excel = excel_files[0]

    # Quét từ file Excel quyết toán chính
    if primary_excel:
        try:
            xl = pd.ExcelFile(primary_excel)
            base_ef = os.path.basename(primary_excel)

            for sname in xl.sheet_names:
                s_lower = sname.lower()
                df = xl.parse(sname)

                # A. Tìm Bảng Nghiệm thu Khối lượng hoàn thành (XNKL)
                if any(k in s_lower for k in ["xnkl", "kl hoàn thành", "kl hoan thanh", "bbnt kl"]):
                    if not res["source_xnkl"]:
                        res["source_xnkl"] = f"{base_ef} (Sheet: {sname})"
                    _extract_xnkl_from_dataframe(df, xnkl_dict)

                # B. Tìm Bảng Quyết toán Vật tư A cấp / Giao nhận hàng hóa
                elif any(k in s_lower for k in ["giao nhận", "giao nhan", "vttb", "vật tư a cấp"]):
                    if not res["source_vttb"]:
                        res["source_vttb"] = f"{base_ef} (Sheet: {sname})"
                    else:
                        res["source_vttb"] += f", {sname}"
                    _extract_vttb_from_dataframe(df, vttb_dict)

                # C. Tìm Bảng Vật tư thu hồi
                elif any(k in s_lower for k in ["thu hồi", "thu hoi", "vtth", "bbdg vtth"]):
                    if not res["source_vtth"]:
                        res["source_vtth"] = f"{base_ef} (Sheet: {sname})"
                    _extract_vtth_from_dataframe(df, vtth_dict)
        except Exception as e:
            res["warnings"].append(f"Lỗi đọc file Excel {os.path.basename(primary_excel)}: {e}")

    # Nếu file chính chưa có bảng VTTH, tìm ở các file Excel khác trong thư mục
    if not vtth_dict:
        for ef in excel_files:
            if ef == primary_excel:
                continue
            try:
                xl = pd.ExcelFile(ef)
                for sname in xl.sheet_names:
                    if any(k in sname.lower() for k in ["thu hồi", "thu hoi", "vtth", "bbdg vtth"]):
                        df = xl.parse(sname)
                        if not res["source_vtth"]:
                            res["source_vtth"] = f"{os.path.basename(ef)} (Sheet: {sname})"
                        _extract_vtth_from_dataframe(df, vtth_dict)
            except:
                pass

    # Fallback: Quét từ PDF nếu chưa có dữ liệu từ Excel
    if not xnkl_dict or not vttb_dict:
        if pdfplumber is not None:
            for pf in pdf_files:
                base_pf = os.path.basename(pf)
                try:
                    with pdfplumber.open(pf) as pdf:
                        for idx, page in enumerate(pdf.pages):
                            tables = page.extract_tables() or []
                            for tbl in tables:
                                if not tbl or len(tbl) < 2:
                                    continue
                                header_str = " ".join([str(c or "") for c in tbl[0]]).lower()
                                if any(k in header_str for k in ["danh mục hàng hóa", "tên vật tư", "nội dung công việc"]):
                                    df_tbl = pd.DataFrame(tbl[1:], columns=tbl[0])
                                    if "nghiệm thu" in header_str or "thực hiện" in header_str:
                                        if not res["source_xnkl"]:
                                            res["source_xnkl"] = f"{base_pf} (Trang {idx+1})"
                                        _extract_xnkl_from_dataframe(df_tbl, xnkl_dict)
                                    if "giao nhận" in header_str or "xuất kho" in header_str or "số lượng" in header_str:
                                        if not res["source_vttb"]:
                                            res["source_vttb"] = f"{base_pf} (Trang {idx+1})"
                                        _extract_vttb_from_dataframe(df_tbl, vttb_dict)
                except Exception as e:
                    res["warnings"].append(f"Lỗi đọc file PDF {base_pf}: {e}")

    # Lập bảng đối chiếu chéo tổng hợp
    all_std_names = sorted(list(set(list(xnkl_dict.keys()) + list(vttb_dict.keys()) + list(vtth_dict.keys()))))
    if all_std_names:
        res["found"] = True

    items = []
    matched_c = 0
    surplus_c = 0
    deficit_c = 0

    for name in all_std_names:
        info_xn = xnkl_dict.get(name, {})
        info_vt = vttb_dict.get(name, {})
        info_th = vtth_dict.get(name, {})

        raw_name = info_vt.get("raw_name") or info_xn.get("raw_name") or info_th.get("raw_name") or name
        dvt = info_vt.get("dvt") or info_xn.get("dvt") or info_th.get("dvt") or "Cái"

        sl_xuat = float(info_vt.get("sl_xuat", 0))
        sl_qt = float(info_vt.get("sl_qt", sl_xuat))  # Nếu chưa có cột riêng thì lấy bằng sl xuất
        sl_thua = float(info_th.get("sl_thu_hoi", 0))
        sl_nt = float(info_xn.get("sl_nt", 0))

        if sl_xuat == 0 and sl_qt == 0 and sl_thua == 0 and sl_nt == 0:
            continue

        chenh_lech = round(sl_qt - sl_nt, 3)

        if abs(chenh_lech) < 0.001:
            ket_luan = "Khớp"
            badge = "🟢"
            status_code = "KHOP"
            matched_c += 1
        elif chenh_lech > 0:
            ket_luan = f"Lệch thừa ({chenh_lech:g})"
            badge = "🔴"
            status_code = "LECH_THUA"
            surplus_c += 1
        else:
            ket_luan = f"Lệch thiếu ({abs(chenh_lech):g})"
            badge = "🟡"
            status_code = "LECH_THIEU"
            deficit_c += 1

        items.append({
            "ten_chuan": name,
            "ten_goc": raw_name,
            "dvt": dvt,
            "sl_xuat": sl_xuat,
            "sl_qt": sl_qt,
            "sl_thua": sl_thua,
            "sl_nt": sl_nt,
            "chenh_lech": chenh_lech,
            "ket_luan": ket_luan,
            "badge": badge,
            "status_code": status_code
        })

    res["items"] = items
    res["matched_count"] = matched_c
    res["surplus_count"] = surplus_c
    res["deficit_count"] = deficit_c
    return res


def _find_table_header_and_cols(df, mode="xnkl"):
    """
    Tìm chính xác dòng bắt đầu dữ liệu, cột tên vật tư, cột ĐVT và cột số lượng.
    mode: "xnkl" (nghiệm thu thực hiện), "vttb" (giao nhận/xuất kho), "vtth" (thu hồi)
    """
    for r_idx in range(min(30, len(df))):
        row = [str(c or "").strip() for c in df.iloc[r_idx]]
        row_lower = [c.lower() for c in row]
        if any(len(c) > 60 for c in row):
            continue

        has_stt = any(re.match(r"^(stt|tt)$", c) for c in row_lower)
        has_name = any(k in c for c in row_lower for k in ["danh mục", "tên vật tư", "tên hàng", "nội dung", "quy cách", "vttb thu hồi"])
        has_sl = any(k in c for c in row_lower for k in ["số lượng", "khối lượng", "thực hiện", "theo bb", "sl giao", "sl xuất"])
        has_dvt = any(k in c for c in row_lower for k in ["đơn vị tính", "đvt", "dvt", "đơn vị"])

        if (has_name and (has_sl or has_dvt)) or (has_stt and has_name):
            name_c = -1
            dvt_c = -1
            sl_c = -1
            for c_idx, c in enumerate(row_lower):
                if any(k in c for k in ["danh mục", "tên vật tư", "tên hàng", "nội dung", "quy cách", "vttb thu hồi"]) and not any(k in c for k in ["biên bản", "bảng giá", "giá trị"]):
                    name_c = c_idx
                elif any(k in c for k in ["đơn vị tính", "đvt", "dvt", "đơn vị"]):
                    dvt_c = c_idx
                elif any(k in c for k in ["số lượng", "khối lượng", "thực hiện", "sl giao", "sl xuất", "theo bb"]):
                    if "%" not in c and "tỉ lệ" not in c and "giá trị" not in c:
                        sl_c = c_idx

            start_row = r_idx + 1
            if r_idx + 1 < len(df):
                sub_row = [str(c or "").strip().lower() for c in df.iloc[r_idx + 1]]
                for sc_idx, sc in enumerate(sub_row):
                    if mode == "xnkl" and "thực hiện" in sc and "%" not in sc and "tỉ lệ" not in sc:
                        sl_c = sc_idx
                    elif mode in ["vttb", "vtth"] and any(k in sc for k in ["số lượng", "thực hiện", "khối lượng"]) and "%" not in sc:
                        sl_c = sc_idx
                if any(k in sub_row for k in ["hợp đồng", "thực hiện", "phát sinh"]):
                    start_row = r_idx + 2

            if name_c != -1:
                return start_row, name_c, dvt_c, sl_c
    return None


def _clean_numeric(val):
    if val is None or pd.isna(val):
        return 0.0
    try:
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip().replace(".", "").replace(",", ".")
        m = re.search(r"[-+]?\d*\.?\d+", s)
        return float(m.group(0)) if m else 0.0
    except:
        return 0.0


def _extract_xnkl_from_dataframe(df, out_dict):
    """Bóc tách vật tư và khối lượng từ Bảng Nghiệm thu hoàn thành (XNKL)."""
    tbl_info = _find_table_header_and_cols(df, mode="xnkl")
    if not tbl_info:
        return
    start_row, name_col, dvt_col, sl_col = tbl_info

    for r in range(start_row, len(df)):
        raw_name = str(df.iloc[r, name_col] or "").strip()
        if not raw_name or raw_name.lower() in ["nan", "tổng cộng", "stt", "hàng hóa", "cộng"]:
            continue
        sl_val = _clean_numeric(df.iloc[r, sl_col]) if sl_col != -1 and sl_col < len(df.columns) else 0.0
        dvt_val = str(df.iloc[r, dvt_col]).strip() if dvt_col != -1 and dvt_col < len(df.columns) and pd.notna(df.iloc[r, dvt_col]) else "Cái"

        # Bỏ qua nếu tên là số hiệu cột hoặc số thứ tự (ví dụ '1.1', '(1)', '(2) (3)')
        if re.match(r"^[\(\)\d\.\s\-]+$", raw_name):
            continue

        # Bỏ qua các dòng dịch vụ, nhân công xây lắp, bảo hiểm không thuộc VTTB
        name_l = raw_name.lower()
        if any(k in name_l for k in ["bảo dưỡng đầu phát", "thi công sửa chữa", "mua bảo hiểm", "nhân công", "chi phí khác", "thí nghiệm", "khảo sát hiện trạng"]):
            continue
        if dvt_val.lower() in ["công trình", "toàn bộ", "hợp đồng", "gói thầu"]:
            continue

        if sl_val > 0 or any(c in raw_name.lower() for c in ["fco", "lbfco", "la", "xi-lanh", "xéc-măng", "bạc"]):
            std_name = normalize_material_name(raw_name)
            if std_name in out_dict:
                out_dict[std_name]["sl_nt"] += sl_val
            else:
                out_dict[std_name] = {
                    "raw_name": raw_name,
                    "dvt": dvt_val,
                    "sl_nt": sl_val
                }


def _extract_vttb_from_dataframe(df, out_dict):
    """Bóc tách vật tư A cấp / Giao nhận hàng hóa."""
    tbl_info = _find_table_header_and_cols(df, mode="vttb")
    if not tbl_info:
        return
    start_row, name_col, dvt_col, sl_col = tbl_info

    for r in range(start_row, len(df)):
        raw_name = str(df.iloc[r, name_col] or "").strip()
        if not raw_name or raw_name.lower() in ["nan", "tổng cộng", "stt", "hàng hóa", "cộng"]:
            continue
        sl_val = _clean_numeric(df.iloc[r, sl_col]) if sl_col != -1 and sl_col < len(df.columns) else 0.0
        dvt_val = str(df.iloc[r, dvt_col]).strip() if dvt_col != -1 and dvt_col < len(df.columns) and pd.notna(df.iloc[r, dvt_col]) else "Cái"

        if re.match(r"^[\(\)\d\.\s\-]+$", raw_name):
            continue

        if sl_val > 0 or any(c in raw_name.lower() for c in ["fco", "lbfco", "la", "xi-lanh", "xéc-măng", "bạc"]):
            std_name = normalize_material_name(raw_name)
            if std_name in out_dict:
                out_dict[std_name]["sl_xuat"] += sl_val
                out_dict[std_name]["sl_qt"] += sl_val
            else:
                out_dict[std_name] = {
                    "raw_name": raw_name,
                    "dvt": dvt_val,
                    "sl_xuat": sl_val,
                    "sl_qt": sl_val
                }


def _extract_vtth_from_dataframe(df, out_dict):
    """Bóc tách vật tư thu hồi thừa nhập kho."""
    tbl_info = _find_table_header_and_cols(df, mode="vtth")
    if not tbl_info:
        return
    start_row, name_col, dvt_col, sl_col = tbl_info
    # Đặc biệt nếu mode vtth và có cột số lượng theo BB
    if sl_col == -1 or sl_col >= len(df.columns):
        sl_col = 4 if len(df.columns) > 4 else 3

    for r in range(start_row, len(df)):
        raw_name = str(df.iloc[r, name_col] or "").strip()
        if not raw_name or raw_name.lower() in ["nan", "tổng cộng", "phần vttb"]:
            continue
        if re.match(r"^[\(\)\d\.\s\-]+$", raw_name):
            continue

        sl_val = _clean_numeric(df.iloc[r, sl_col]) if sl_col != -1 and sl_col < len(df.columns) else 0.0
        dvt_val = str(df.iloc[r, dvt_col]).strip() if dvt_col != -1 and dvt_col < len(df.columns) and pd.notna(df.iloc[r, dvt_col]) else "Bộ"

        if sl_val > 0:
            std_name = normalize_material_name(raw_name)
            if std_name in out_dict:
                out_dict[std_name]["sl_thu_hoi"] += sl_val
            else:
                out_dict[std_name] = {
                    "raw_name": raw_name,
                    "dvt": dvt_val,
                    "sl_thu_hoi": sl_val
                }


# ── 5. XUẤT BÁO CÁO EXCEL CHUẨN EVN (Times New Roman 12, 4 Sheet) ──
def export_excel_qd202_report(ma_ct, ten_ct, audit_data, recon_data):
    """
    Tạo file Excel BaoCao_KiemTra_SCL_QD202.xlsx định dạng Times New Roman 12, số #,##0,
    gồm 4 sheet chuẩn theo đúng yêu cầu:
    1_TongHop_QD202
    2_DoiChieu_VatTu_XNKL
    3_ChiTiet_HoSo
    4_CanhBao_RuiRo
    """
    wb = openpyxl.Workbook()
    font_tnr_12 = Font(name="Times New Roman", size=12)
    font_tnr_12_bold = Font(name="Times New Roman", size=12, bold=True)
    font_title = Font(name="Times New Roman", size=14, bold=True, color="1565C0")
    fill_header = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    fill_danger = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    fill_ok = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
    border_thin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    # ── SHEET 1: 1_TongHop_QD202 ──
    ws1 = wb.active
    ws1.title = "1_TongHop_QD202"

    ws1["A1"] = f"BÁO CÁO TỔNG HỢP KIỂM TRA HỒ SƠ QUYẾT TOÁN SCL THEO QĐ 202/QĐ-HĐTV"
    ws1["A1"].font = font_title
    ws1["A2"] = f"Công trình: {ma_ct} - {ten_ct}"
    ws1["A2"].font = font_tnr_12_bold
    ws1["A3"] = f"Ngày kiểm tra: {datetime.date.today().strftime('%d/%m/%Y')} | Hình thức: {'Tự làm' if audit_data.get('is_tu_lam') else 'Thuê ngoài'}"
    ws1["A3"].font = font_tnr_12

    headers_ws1 = ["STT", "Nhóm Hồ Sơ QĐ 202", "Mô tả nội dung", "Số file", "Số file Scan (chưa OCR)", "Tình trạng Tuân thủ"]
    for c_idx, h in enumerate(headers_ws1, 1):
        cell = ws1.cell(5, c_idx, h)
        cell.font = font_tnr_12_bold
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin

    row_idx = 6
    for g_id, g_info in QD202_GROUPS.items():
        g_files = audit_data["groups"].get(g_id, [])
        scanned_count = sum(1 for f in g_files if f["is_scanned"])
        status_txt = "Đầy đủ" if len(g_files) >= len(g_info["required_docs"]) - 1 else f"Thiếu ({len(g_info['required_docs']) - len(g_files)} mục)"
        if g_id == 2 and audit_data.get("is_tu_lam"):
            status_txt = "Áp dụng theo Thỏa thuận giao việc (TTGV)"

        ws1.cell(row_idx, 1, g_id).alignment = Alignment(horizontal="center")
        ws1.cell(row_idx, 2, g_info["name"]).font = font_tnr_12_bold
        ws1.cell(row_idx, 3, g_info["desc"]).font = font_tnr_12
        ws1.cell(row_idx, 4, len(g_files)).number_format = "#,##0"
        ws1.cell(row_idx, 5, scanned_count).number_format = "#,##0"
        cell_st = ws1.cell(row_idx, 6, status_txt)
        cell_st.font = font_tnr_12_bold
        if "Thiếu" in status_txt:
            cell_st.fill = fill_warn
        else:
            cell_st.fill = fill_ok

        for col in range(1, 7):
            ws1.cell(row_idx, col).border = border_thin
            if col not in [2, 6]:
                ws1.cell(row_idx, col).font = font_tnr_12
        row_idx += 1

    # ── SHEET 2: 2_DoiChieu_VatTu_XNKL ──
    ws2 = wb.create_sheet("2_DoiChieu_VatTu_XNKL")
    ws2["A1"] = f"BẢNG ĐỐI CHIẾU CHÉO VẬT TƯ A CẤP VÀ NGHIỆM THU KHỐI LƯỢNG (XNKL)"
    ws2["A1"].font = font_title
    ws2["A2"] = f"Nguồn đối chiếu XNKL: {recon_data.get('source_xnkl', 'Chưa có')} | Nguồn Quyết toán: {recon_data.get('source_vttb', 'Chưa có')}"
    ws2["A2"].font = font_tnr_12

    headers_ws2 = [
        "STT", "Tên Vật tư (Chuẩn hóa)", "Tên gốc trong hồ sơ", "ĐVT",
        "SL Xuất kho giao ĐVTC", "SL Quyết toán vật tư A cấp", "SL Vật tư thừa thu hồi",
        "SL Nghiệm thu trên Bảng XNKL", "Chênh lệch (QT Vật tư - XNKL)", "Kết luận kiểm tra"
    ]
    for c_idx, h in enumerate(headers_ws2, 1):
        cell = ws2.cell(4, c_idx, h)
        cell.font = font_tnr_12_bold
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin

    row_idx = 5
    for idx, item in enumerate(recon_data.get("items", []), 1):
        ws2.cell(row_idx, 1, idx).alignment = Alignment(horizontal="center")
        ws2.cell(row_idx, 2, item["ten_chuan"]).font = font_tnr_12_bold
        ws2.cell(row_idx, 3, item["ten_goc"]).font = font_tnr_12
        ws2.cell(row_idx, 4, item["dvt"]).alignment = Alignment(horizontal="center")

        ws2.cell(row_idx, 5, item["sl_xuat"]).number_format = "#,##0"
        ws2.cell(row_idx, 6, item["sl_qt"]).number_format = "#,##0"
        ws2.cell(row_idx, 7, item["sl_thua"]).number_format = "#,##0"
        ws2.cell(row_idx, 8, item["sl_nt"]).number_format = "#,##0"

        cell_diff = ws2.cell(row_idx, 9, item["chenh_lech"])
        cell_diff.number_format = "#,##0"
        cell_diff.font = font_tnr_12_bold

        cell_kl = ws2.cell(row_idx, 10, item["ket_luan"])
        cell_kl.font = font_tnr_12_bold
        if item["status_code"] == "LECH_THUA":
            cell_kl.fill = fill_danger
        elif item["status_code"] == "LECH_THIEU":
            cell_kl.fill = fill_warn
        else:
            cell_kl.fill = fill_ok

        for col in range(1, 11):
            ws2.cell(row_idx, col).border = border_thin
            if col not in [2, 9, 10]:
                ws2.cell(row_idx, col).font = font_tnr_12
        row_idx += 1

    # ── SHEET 3: 3_ChiTiet_HoSo ──
    ws3 = wb.create_sheet("3_ChiTiet_HoSo")
    ws3["A1"] = f"DANH MỤC CHI TIẾT FILE VÀ PHÂN LOẠI THEO QUYẾT ĐỊNH 202"
    ws3["A1"].font = font_title

    headers_ws3 = ["STT", "Tên file tài liệu", "Nhóm QĐ 202", "Loại chứng từ", "Số hiệu QĐ / HĐ", "Ngày văn bản", "Định dạng", "Cảnh báo OCR Scan"]
    for c_idx, h in enumerate(headers_ws3, 1):
        cell = ws3.cell(3, c_idx, h)
        cell.font = font_tnr_12_bold
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin

    row_idx = 4
    for idx, f in enumerate(audit_data.get("all_files", []), 1):
        ws3.cell(row_idx, 1, idx).alignment = Alignment(horizontal="center")
        ws3.cell(row_idx, 2, f["name"]).font = font_tnr_12
        g_name = QD202_GROUPS.get(f["group_id"], {}).get("name", "Chưa phân loại")
        ws3.cell(row_idx, 3, g_name).font = font_tnr_12
        ws3.cell(row_idx, 4, f["doc_name"]).font = font_tnr_12
        ws3.cell(row_idx, 5, f["so_qd"]).font = font_tnr_12
        ws3.cell(row_idx, 6, f["ngay_str"]).alignment = Alignment(horizontal="center")
        ws3.cell(row_idx, 7, f["ext"].upper()).alignment = Alignment(horizontal="center")

        cell_ocr = ws3.cell(row_idx, 8, "⚠️ Scan ảnh (Chưa OCR)" if f["is_scanned"] else "✅ Text")
        cell_ocr.font = font_tnr_12_bold
        if f["is_scanned"]:
            cell_ocr.fill = fill_warn

        for col in range(1, 9):
            ws3.cell(row_idx, col).border = border_thin
            if col != 8:
                ws3.cell(row_idx, col).font = font_tnr_12
        row_idx += 1

    # ── SHEET 4: 4_CanhBao_RuiRo ──
    ws4 = wb.create_sheet("4_CanhBao_RuiRo")
    ws4["A1"] = f"TỔNG HỢP CẢNH BÁO RỦI RO & DANH MỤC HỒ SƠ CẦN BỔ SUNG"
    ws4["A1"].font = font_title

    headers_ws4 = ["STT", "Mức độ", "Nhóm / Hạng mục", "Nội dung cảnh báo sai sót", "Hành động khắc phục yêu cầu"]
    for c_idx, h in enumerate(headers_ws4, 1):
        cell = ws4.cell(3, c_idx, h)
        cell.font = font_tnr_12_bold
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin

    row_idx = 4
    # Danh mục hồ sơ thiếu
    for idx, m in enumerate(audit_data.get("missing_docs", []), 1):
        ws4.cell(row_idx, 1, idx).alignment = Alignment(horizontal="center")
        cell_sev = ws4.cell(row_idx, 2, "NGHIÊM TRỌNG" if m["severity"] == "DANGER" else "CẦN LƯU Ý")
        cell_sev.font = font_tnr_12_bold
        cell_sev.fill = fill_danger if m["severity"] == "DANGER" else fill_warn

        ws4.cell(row_idx, 3, m["group_name"]).font = font_tnr_12
        ws4.cell(row_idx, 4, f"Thiếu chứng từ: {m['doc_name']}").font = font_tnr_12
        ws4.cell(row_idx, 5, f"Thu thập và bổ sung file {m['doc_name']} vào hồ sơ công trình trước khi thẩm tra.").font = font_tnr_12

        for col in range(1, 6):
            ws4.cell(row_idx, col).border = border_thin
        row_idx += 1

    # Cảnh báo rủi ro thời gian
    for tw in audit_data.get("timeline_warnings", []):
        ws4.cell(row_idx, 1, row_idx - 3).alignment = Alignment(horizontal="center")
        cell_sev = ws4.cell(row_idx, 2, "NGHIÊM TRỌNG")
        cell_sev.font = font_tnr_12_bold
        cell_sev.fill = fill_danger

        ws4.cell(row_idx, 3, "Trình tự thời gian").font = font_tnr_12
        ws4.cell(row_idx, 4, tw).font = font_tnr_12
        ws4.cell(row_idx, 5, "Rà soát lại ngày ký các biên bản hoặc bổ sung Phụ lục gia hạn hợp đồng phù hợp.").font = font_tnr_12

        for col in range(1, 6):
            ws4.cell(row_idx, col).border = border_thin
        row_idx += 1

    # Cảnh báo chênh lệch vật tư
    for item in recon_data.get("items", []):
        if item["status_code"] in ["LECH_THUA", "LECH_THIEU"]:
            ws4.cell(row_idx, 1, row_idx - 3).alignment = Alignment(horizontal="center")
            cell_sev = ws4.cell(row_idx, 2, "CẦN LÀM RÕ")
            cell_sev.font = font_tnr_12_bold
            cell_sev.fill = fill_warn

            ws4.cell(row_idx, 3, "Đối chiếu vật tư").font = font_tnr_12
            ws4.cell(row_idx, 4, f"Vật tư '{item['ten_chuan']}': {item['ket_luan']} (QT: {item['sl_qt']:g} {item['dvt']}, XNKL: {item['sl_nt']:g} {item['dvt']})").font = font_tnr_12
            ws4.cell(row_idx, 5, "Yêu cầu ĐVTC và Đội quản lý giải trình lý do chênh lệch hoặc điều chỉnh lại Bảng XNKL / Quyết toán.").font = font_tnr_12

            for col in range(1, 6):
                ws4.cell(row_idx, col).border = border_thin
            row_idx += 1

    # ── Auto fit column widths ──
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if "\n" in val_str:
                    lines = val_str.split("\n")
                    max_len = max(max_len, max(len(l) for l in lines))
                else:
                    max_len = max(max_len, len(val_str))
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
