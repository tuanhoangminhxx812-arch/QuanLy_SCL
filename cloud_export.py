"""
Module xuất file Word cho app Cloud (Read-only).
Hỗ trợ 4 mẫu: TMQT, Phiếu thẩm tra, Báo cáo thẩm tra, QĐ phê duyệt.
"""
import os, datetime, re
import pandas as pd
from io import BytesIO
from form_module import doc_so_vn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _safe_int(v):
    try:
        return int(float(v)) if pd.notna(v) else 0
    except:
        return 0


def _fmt_money_dot(val):
    if val == 0:
        return '0'
    return f'{val:,}'.replace(',', '.')


def _format_date_vn(d):
    if pd.isna(d) or d is None:
        return '....../....../...........'
    if isinstance(d, pd.Timestamp):
        d = d.date()
    if isinstance(d, (datetime.date, datetime.datetime)):
        return d.strftime('%d/%m/%Y')
    return str(d)


def get_project_section(db_df, ten_ct, ma_ct=''):
    """Tìm section của công trình trong database, trả về (main_row, ct_data)."""
    start_indices = db_df.index[db_df['Tên Công trình'] == ten_ct].tolist()
    if not start_indices:
        start_indices = db_df.index[db_df['Mã CT'].astype(str).str.strip() == str(ma_ct).strip()].tolist()
    if not start_indices:
        return None, pd.DataFrame()
    start_idx = start_indices[0]
    end_idx = len(db_df)
    for i in range(start_idx + 1, len(db_df)):
        val = str(db_df.at[i, 'STT']).strip().upper()
        if val in ['I','II','III','IV','V','VI','VII','VIII','IX','X']:
            end_idx = i
            break
    ct_data = db_df.iloc[start_idx:end_idx]
    return ct_data.iloc[0], ct_data


def get_cost_breakdown(ct_data):
    """Trích xuất chi phí A/B/C/D/SCL từ section của công trình."""
    result = {}
    for _, row in ct_data.iterrows():
        stt = str(row['STT']).strip()
        if stt in ['A', 'B', 'C', 'D', 'E', 'F', 'SCL']:
            dt = _safe_int(row.get('Giá trị Dự toán', 0))
            qt = _safe_int(row.get('Giá trị Q.định phê duyệt QT công trình', 0))
            if stt not in result or qt != 0:
                result[stt] = {'dt': dt, 'qt': qt}
    return result


def _replace_para_lines(p, lines):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    if not lines:
        p.text = ""
        return
    style = p.style
    li = p.paragraph_format.left_indent
    fli = p.paragraph_format.first_line_indent
    for line in lines[:-1]:
        np = p.insert_paragraph_before(line, style=style)
        np.paragraph_format.left_indent = li
        np.paragraph_format.first_line_indent = fli
        np.paragraph_format.space_after = Pt(0)
        np.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.text = lines[-1]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _apply_font(doc):
    from docx.shared import Pt, Cm
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(0)
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)


# ============================================================
# 1. TMQT - Bản thuyết minh quyết toán
# ============================================================
def export_tmqt_word(main_row, ct_data, noi_dung_scl=''):
    from docx import Document as DocxDocument
    from docx.shared import Pt
    template = os.path.join(BASE_DIR, 'Mẫu TMQT.docx')
    if not os.path.exists(template):
        return None
    doc = DocxDocument(template)
    now = datetime.datetime.now()

    ten = str(main_row.get('Tên Công trình', ''))
    ma = str(main_row.get('Mã CT', '')) if pd.notna(main_row.get('Mã CT')) else ''
    ke_hoach = _safe_int(main_row.get('Kế hoạch', 0))
    don_vi = str(main_row.get('Đơn vị QL', '')) if pd.notna(main_row.get('Đơn vị QL')) else ''
    ghi_chu = str(main_row.get('Ghi chú', '')) if pd.notna(main_row.get('Ghi chú')) else ''
    can_cu = str(main_row.get('Căn cứ pháp lý', '')) if pd.notna(main_row.get('Căn cứ pháp lý')) else ''
    klcv = str(main_row.get('Khối lượng công việc', '')) if pd.notna(main_row.get('Khối lượng công việc')) else ''
    if not klcv and noi_dung_scl:
        klcv = noi_dung_scl
    kc = _format_date_vn(main_row.get('Ngày khởi công'))
    ht = _format_date_vn(main_row.get('Ngày hoàn thành'))

    bd = get_cost_breakdown(ct_data)
    gt_dt = bd.get('SCL', {}).get('dt', 0)
    gt_qt = bd.get('SCL', {}).get('qt', 0)

    for p in doc.paragraphs:
        p.paragraph_format.space_after = Pt(0)
        t = p.text.strip()
        if not t:
            continue
        if "- Tên danh mục:" in t:
            _replace_para_lines(p, [f"- Tên danh mục: {ten}", f"- Mã công trình: {ma}"])
        elif "- Giá trị vốn kế hoạch:" in t:
            p.text = f"- Giá trị vốn kế hoạch: {_fmt_money_dot(ke_hoach)} đồng"
        elif "sửa chữa lớn năm" in t:
            p.text = f"- Thuộc kế hoạch vốn sửa chữa lớn năm {now.year}"
        elif "Hình thức tự làm hay thuê ngoài" in t:
            p.text = f"- Hình thức tự làm hay thuê ngoài: {ghi_chu}"
        elif "- Tên đơn vị thi công" in t:
            p.text = f"- Tên đơn vị thi công: {don_vi}"
        elif "- Giá trị dự toán được duyệt" in t:
            p.text = f"- Giá trị dự toán được duyệt: {_fmt_money_dot(gt_dt)} đồng"
        elif "- Thời gian khởi công" in t:
            p.text = f"- Thời gian khởi công: {kc}"
        elif "- Thời gian hoàn thành" in t:
            p.text = f"- Thời gian hoàn thành: {ht}"
        elif "- Giá trị quyết toán" in t and "hoàn thành" in t:
            p.text = f"- Giá trị quyết toán danh mục hoàn thành: {_fmt_money_dot(gt_qt)} đồng"
        elif "Khối lượng công việc chủ yếu đã tiến hành" in t:
            lines = ["- Khối lượng công việc chủ yếu đã tiến hành (thay thế, sửa chữa những bộ phận nào của TSCĐ):"]
            if klcv:
                lines.extend([l for l in klcv.split('\n') if l.strip()])
            _replace_para_lines(p, lines)
        elif "Các căn cứ về chế độ để lập quyết toán" in t:
            lines = ["- Các căn cứ về chế độ để lập quyết toán:"]
            if can_cu:
                lines.extend([l for l in can_cu.split('\n') if l.strip()])
            _replace_para_lines(p, lines)
        elif "+ .........." in t:
            p.text = ""
        elif "Phân tích các nhân tố tăng giảm" in t:
            chenh = gt_qt - gt_dt
            if chenh > 0:
                p.text = f"- Phân tích: Giá trị quyết toán tăng {_fmt_money_dot(abs(chenh))} đồng so với dự toán được duyệt."
            elif chenh < 0:
                p.text = f"- Phân tích: Giá trị quyết toán giảm {_fmt_money_dot(abs(chenh))} đồng so với dự toán được duyệt."
            else:
                p.text = "- Phân tích: Giá trị quyết toán bằng dự toán được duyệt."
        elif "ngày       tháng      năm" in t:
            p.text = t.replace("2026", str(now.year))

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# 2. Phiếu thẩm tra quyết toán
# ============================================================
def export_phieu_tham_tra_word(main_row):
    from docx import Document as DocxDocument
    from docx.shared import Pt
    template = os.path.join(BASE_DIR, 'Mẫu phiếu thẩm tra QT.docx')
    if not os.path.exists(template):
        return None
    doc = DocxDocument(template)

    ten = str(main_row.get('Tên Công trình', ''))
    ma = str(main_row.get('Mã CT', '')) if pd.notna(main_row.get('Mã CT')) else ''
    don_vi = str(main_row.get('Đơn vị QL', '')) if pd.notna(main_row.get('Đơn vị QL')) else ''
    ghi_chu = str(main_row.get('Ghi chú', '')) if pd.notna(main_row.get('Ghi chú')) else ''
    so_hd = str(main_row.get('Số Hợp đồng xây lắp', '')) if pd.notna(main_row.get('Số Hợp đồng xây lắp')) else ''
    ngay_hd = _format_date_vn(main_row.get('Ngày Hợp đồng xây lắp'))

    is_tu_lam = 'tự' in ghi_chu.lower() if ghi_chu else False

    for p in doc.paragraphs:
        t = p.text.strip()
        if "Tên công trình SCL:" in t:
            p.text = f"Tên công trình SCL: {ten}"
        elif "Mã công trình:" in t:
            p.text = f"Mã công trình: {ma}"
        elif "Đơn vị quản lý:" in t:
            p.text = f"Đơn vị quản lý: {don_vi}"
        elif "Phương thức chọn thầu" in t:
            if is_tu_lam:
                p.text = "Phương thức chọn thầu thực hiện: Tự làm"
            else:
                p.text = f"Phương thức chọn thầu thực hiện: Thuê ngoài"
        elif t == "Tự làm":
            if is_tu_lam:
                p.text = "✓ Tự làm"
            else:
                p.text = "☐ Tự làm"
        elif "Thuê ngoài (hợp đồng số" in t:
            if not is_tu_lam and so_hd:
                p.text = f"✓ Thuê ngoài (hợp đồng số {so_hd} ngày {ngay_hd})"
            elif not is_tu_lam:
                p.text = "✓ Thuê ngoài (hợp đồng số ......... ngày ......)"
            else:
                p.text = "☐ Thuê ngoài (hợp đồng số ......... ngày ......)"
        elif "Đơn vị thực hiện:" in t:
            p.text = f"Đơn vị thực hiện: {don_vi}"

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# 3. Báo cáo thẩm tra quyết toán
# ============================================================
def export_bao_cao_tham_tra_word(main_row, ct_data):
    from docx import Document as DocxDocument
    from docx.shared import Pt
    template = os.path.join(BASE_DIR, 'Mẫu phiếu báo cáo thẩm tra QT.docx')
    if not os.path.exists(template):
        return None
    doc = DocxDocument(template)

    ten = str(main_row.get('Tên Công trình', ''))
    don_vi = str(main_row.get('Đơn vị QL', '')) if pd.notna(main_row.get('Đơn vị QL')) else ''
    ghi_chu = str(main_row.get('Ghi chú', '')) if pd.notna(main_row.get('Ghi chú')) else ''
    so_hd = str(main_row.get('Số Hợp đồng xây lắp', '')) if pd.notna(main_row.get('Số Hợp đồng xây lắp')) else ''

    bd = get_cost_breakdown(ct_data)
    a_dt = bd.get('A', {}).get('dt', 0)
    a_qt = bd.get('A', {}).get('qt', 0)
    b_dt = bd.get('B', {}).get('dt', 0)
    b_qt = bd.get('B', {}).get('qt', 0)
    c_dt = bd.get('C', {}).get('dt', 0)
    c_qt = bd.get('C', {}).get('qt', 0)
    scl_qt = bd.get('SCL', {}).get('qt', 0)

    paragraphs = doc.paragraphs
    for i, p in enumerate(paragraphs):
        t = p.text.strip()
        if t == "Tên công trình SCL:":
            p.text = f"Tên công trình SCL: {ten}"
        elif t == "Đơn vị quản lý:":
            p.text = f"Đơn vị quản lý: {don_vi}"
        elif "Phương thức chọn thầu" in t:
            is_tu = 'tự' in ghi_chu.lower()
            p.text = f"Phương thức chọn thầu thực hiện: {'Tự làm' if is_tu else 'Thuê ngoài'}"
        elif t == "Hợp đồng số:":
            p.text = f"Hợp đồng số: {so_hd}" if so_hd else "Hợp đồng số: "
        elif t == "Đơn vị thực hiện:":
            p.text = f"Đơn vị thực hiện: {don_vi}"
        elif "1/ Phần xây dựng:" in t:
            p.text = "\t1/ Phần xây dựng:"
        elif "Số dự toán:" in t and "2/" not in t and "3/" not in t:
            if i > 0 and "1/" in str(paragraphs[i-1].text):
                p.text = f"\t- Số dự toán: {_fmt_money_dot(b_dt)} đồng"
            else:
                p.text = f"- Số dự toán: {_fmt_money_dot(a_dt)} đồng"
        elif "Số quyết toán:" in t:
            # Check context
            p.text = f"\t- Số quyết toán: đồng"
        elif "Số thẩm tra:" in t:
            p.text = f"\t- Số thẩm tra: đồng"
        elif "Số chênh lệch:" in t:
            p.text = f"          - Số chênh lệch: đồng"
        elif "tổ thẩm tra quyết toán chấp thuận" in t.lower():
            p.text = f"Sau khi xem xét thẩm tra hồ sơ, tổ thẩm tra quyết toán chấp thuận tổng giá trị quyết toán công trình nêu trên: {_fmt_money_dot(scl_qt)} đồng"
        elif "Tổng số:" in t:
            p.text = f"Tổng số: {_fmt_money_dot(scl_qt)} đồng"
        elif "Xây dựng:" in t and "Phần" not in t:
            p.text = f"Xây dựng: {_fmt_money_dot(b_qt)} đồng"
        elif "Thiết bị:" in t and "Phần" not in t:
            p.text = f"Thiết bị: {_fmt_money_dot(a_qt)} đồng"
        elif "KTCB khác:" in t:
            p.text = f" KTCB khác: {_fmt_money_dot(c_qt)} đồng"

    # Fill table if exists
    if doc.tables:
        table = doc.tables[0]
        # Table: Nội dung | Dự toán | QT | Thẩm tra | Chênh lệch
        # Skip header row 0
        # We don't have detailed breakdown for thiết kế/thẩm định, leave blank

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# 4. Quyết định phê duyệt quyết toán
# ============================================================
def export_qd_phe_duyet_word(main_row, ct_data):
    from docx import Document as DocxDocument
    from docx.shared import Pt
    template = os.path.join(BASE_DIR, 'Mẫu QĐ QTCT.docx')
    if not os.path.exists(template):
        return None
    doc = DocxDocument(template)
    now = datetime.datetime.now()

    ten = str(main_row.get('Tên Công trình', ''))
    can_cu = str(main_row.get('Căn cứ pháp lý', '')) if pd.notna(main_row.get('Căn cứ pháp lý')) else ''
    don_vi = str(main_row.get('Đơn vị QL', '')) if pd.notna(main_row.get('Đơn vị QL')) else 'Vũng Tàu'

    bd = get_cost_breakdown(ct_data)
    a_qt = bd.get('A', {}).get('qt', 0)
    b_qt = bd.get('B', {}).get('qt', 0)
    c_qt = bd.get('C', {}).get('qt', 0)
    scl_qt = bd.get('SCL', {}).get('qt', 0)
    bang_chu = doc_so_vn(scl_qt) if scl_qt > 0 else ''

    for p in doc.paragraphs:
        t = p.text.strip()
        if "Tên công trình:" in t and "Điều" not in t:
            p.text = f"Tên công trình: {ten}"
        elif "Điều 1:" in t:
            p.text = f"Điều 1: Phê duyệt quyết toán công trình: {ten} với tổng giá trị: {_fmt_money_dot(scl_qt)} đồng (Bằng chữ: {bang_chu})"
        elif "Chi phí thiết bị:" in t:
            p.text = f"Chi phí thiết bị: {_fmt_money_dot(a_qt)} đồng"
        elif "Chi phí xây dựng:" in t:
            p.text = f"Chi phí xây dựng: {_fmt_money_dot(b_qt)} đồng"
        elif "KTCB khác:" in t:
            p.text = f"KTCB khác: {_fmt_money_dot(c_qt)} đồng"
        elif "Điều 2:" in t:
            p.text = f"Điều 2: Nguồn vốn thực hiện công trình: Sửa chữa lớn của {don_vi}"
        elif "ngày ….tháng……năm" in t:
            p.text = t.replace("ngày ….tháng……năm…", f"ngày ... tháng ... năm {now.year}")

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
