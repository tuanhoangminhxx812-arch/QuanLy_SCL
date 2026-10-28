"""
cloud_export.py
Module xuất 5 mẫu văn bản quyết toán SCL sang định dạng Word (.docx)
Thiết kế chuẩn xác 100% theo bộ hồ sơ mẫu thực tế đã hoàn thiện tại thư mục Tham khao (PCVT)
và Phụ lục 10 - Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của EVNHCMC:

1. Tờ trình phê duyệt quyết toán danh mục SCL hoàn thành (Mẫu thực tế Tham khao/TTr Duyệt QT DM SCL Hoàn thành.pdf)
2. Bản thuyết minh quyết toán (Mẫu thực tế Tham khao/Thuyết minh quyết toán.pdf)
3. Phiếu thẩm tra quyết toán (Mẫu Phụ lục 10.4 QĐ 202/QĐ-HĐTV)
4. Báo cáo kết quả thẩm tra quyết toán (Mẫu thực tế Tham khao/BC Tham Tra QT.pdf)
5. Quyết định phê duyệt quyết toán (Mẫu thực tế Tham khao/525.Quyết định phê duyệt quyết toán...)
"""

import os
import re
import datetime
from io import BytesIO
import pandas as pd
from docx import Document as DocxDocument
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from form_module import doc_so_vn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _safe_int(val):
    try:
        return int(round(float(val))) if pd.notna(val) else 0
    except:
        return 0


def _fmt(val):
    if pd.isna(val) or val == 0:
        return '-'
    v = int(round(val))
    return f"{v:,}".replace(',', '.')


def _fmt_money_dot(val):
    v = int(round(val)) if pd.notna(val) else 0
    return f"{v:,}".replace(',', '.')


def _format_date_vn(d):
    if pd.isna(d) or d is None:
        return '....../....../...........'
    if isinstance(d, pd.Timestamp):
        d = d.date()
    if isinstance(d, (datetime.date, datetime.datetime)):
        return d.strftime('%d/%m/%Y')
    return str(d)


def add_aligned_cost_lines_docx(doc, items, left_pt=49.65, colon_pt=205.55, right_pt=382.75):
    """
    Thêm các dòng số liệu chi phí ngay hàng thẳng lối 100% chuẩn theo mẫu HS_mau
    sử dụng chính xác 3 Tab Stops của Word:
    - Tab 1: lề bắt đầu nhãn
    - Tab 2: vị trí dấu hai chấm (:)
    - Tab 3: vị trí kết thúc số tiền (căn phải)
    """
    for item in items:
        lbl, val, is_bold = item[0], item[1], item[2]
        unit = item[3] if len(item) > 3 else "đồng."
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.add_tab_stop(Pt(left_pt), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Pt(colon_pt), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Pt(right_pt), WD_TAB_ALIGNMENT.RIGHT)
        
        r1 = p.add_run(f"\t{lbl}")
        r2 = p.add_run("\t:  \t")
        r3 = p.add_run(f"{val} {unit.strip()}")
        if is_bold:
            r1.bold = True
            r2.bold = True
            r3.bold = True


def _apply_font(doc, font_name='Times New Roman', font_size=Pt(12)):
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = font_name
            if run.font.size is None:
                run.font.size = font_size
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.name = font_name
                        if run.font.size is None:
                            run.font.size = Pt(11)


def get_cost_breakdown(ct_data):
    """Trích xuất số liệu chi phí từ ct_data trong database_cong_trinh.xlsx."""
    bd = {}
    if ct_data is None or ct_data.empty:
        return bd
    for _, r in ct_data.iterrows():
        stt = str(r.get('STT', '')).strip()
        if not stt:
            continue
        dt = float(r.get('Giá trị Dự toán', 0)) if pd.notna(r.get('Giá trị Dự toán')) else 0.0
        qt = float(r.get('Giá trị Q.định phê duyệt QT công trình', 0)) if pd.notna(r.get('Giá trị Q.định phê duyệt QT công trình')) else 0.0
        cl = dt - qt
        name = str(r.get('Tên Công trình', '')).strip() if pd.notna(r.get('Tên Công trình')) else ''
        bd[stt] = {'dt': dt, 'qt': qt, 'cl': cl, 'name': name}
    return bd


def get_project_section(db_df, project_name, project_code=None):
    """Lấy main_row và toàn bộ ct_data của 1 công trình từ database_cong_trinh.xlsx."""
    if db_df.empty:
        return None, pd.DataFrame()
    p_code = str(project_code).strip().lower() if project_code else None
    p_name = str(project_name).strip().lower() if project_name else None

    main_idx = None
    for idx, row in db_df.iterrows():
        stt = str(row.get('STT', '')).strip()
        ma = str(row.get('Mã CT', '')).strip().lower()
        ten = str(row.get('Tên Công trình', '')).strip().lower()

        if p_code and ma == p_code and stt in ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', '1', '2', '3']:
            main_idx = idx
            break
        elif p_name and p_name in ten and stt in ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', '1', '2', '3']:
            main_idx = idx
            break

    if main_idx is None:
        if p_code:
            match = db_df[db_df['Mã CT'].astype(str).str.strip().str.lower() == p_code]
            if not match.empty:
                main_idx = match.index[0]

    if main_idx is None:
        return None, pd.DataFrame()

    main_row = db_df.loc[main_idx]
    ma_target = str(main_row.get('Mã CT', '')).strip()
    match_all = db_df[db_df['Mã CT'].astype(str).str.strip() == ma_target]
    if len(match_all) > 1:
        return main_row, match_all

    end_idx = len(db_df)
    for i in range(main_idx + 1, len(db_df)):
        stt_next = str(db_df.loc[i, 'STT']).strip()
        if stt_next in ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']:
            end_idx = i
            break
    ct_data = db_df.loc[main_idx:end_idx - 1]
    return main_row, ct_data


# ============================================================
# 1. TỜ TRÌNH PHÊ DUYỆT QUYẾT TOÁN DANH MỤC SCL HOÀN THÀNH
# (Khớp y chang file Tham khao/TTr Duyệt QT DM SCL Hoàn thành.pdf)
# ============================================================
# ============================================================
# 1. TỜ TRÌNH PHÊ DUYỆT QUYẾT TOÁN DANH MỤC SCL HOÀN THÀNH
# (Khớp y chang file HS_mau/TTr Duyệt QT DM SCL Hoàn thành.docx)
# ============================================================
def export_ttr_duyet_qt_word(main_row, ct_data, thqt_data=None):
    doc = DocxDocument()
    
    # Thiết lập lề trang
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(0.8)

    ten = str(main_row.get('Tên Công trình', '')) if main_row is not None else ''
    ma = str(main_row.get('Mã CT', '')) if main_row is not None and pd.notna(main_row.get('Mã CT')) else ''
    so_dt = str(main_row.get('Số Dự toán', '')) if main_row is not None and pd.notna(main_row.get('Số Dự toán')) else '135/QĐ-PCVT'
    ngay_dt = _format_date_vn(main_row.get('Ngày Dự toán')) if main_row is not None else '16/03/2026'
    if '....' in ngay_dt:
        ngay_dt = '16/03/2026'

    # Lấy số liệu từ thqt_data nếu có, ngược lại fallback ct_data
    if thqt_data and 'summary' in thqt_data:
        sm = thqt_data['summary']
        if thqt_data.get('ten_ct'): ten = thqt_data['ten_ct']
        if thqt_data.get('ma_ct'): ma = thqt_data['ma_ct']
        sc_qt = sm.get('sc_qt', 0)
        tb_qt = sm.get('tb_qt', 0)
        khac_qt = sm.get('khac_qt', 0)
        truoc_thue = sm.get('truoc_thue_qt', 0)
        thue_qt = sm.get('thue_qt', 0)
        tong_qt = sm.get('tong_qt', 0)
        vtth_qt = sm.get('vtth_qt', 0)
        scl_qt = sm.get('scl_qt', 0)
        bang_chu = sm.get('bang_chu') or doc_so_vn(tong_qt)
    else:
        bd = get_cost_breakdown(ct_data)
        scl_qt = bd.get('SCL', {}).get('qt', 0)
        if scl_qt == 0 and main_row is not None and 'Giá trị Q.định phê duyệt QT công trình' in main_row:
            scl_qt = float(main_row.get('Giá trị Q.định phê duyệt QT công trình', 0))
        thue_qt = bd.get('B.8', {}).get('qt', 0)
        sc_qt = bd.get('B', {}).get('qt', 0)
        tb_qt = bd.get('A', {}).get('qt', 0)
        khac_qt = bd.get('C', {}).get('qt', 0)
        vtth_qt = bd.get('F', {}).get('qt', 0)
        if thue_qt == 0 and scl_qt > 0:
            thue_qt = int(round(scl_qt - scl_qt / 1.08))
        truoc_thue = scl_qt - thue_qt if scl_qt >= thue_qt else scl_qt
        tong_qt = scl_qt
        bang_chu = doc_so_vn(tong_qt)

    # Bảng tiêu ngữ (4 rows x 2 cols giống HS_mau)
    table_hdr = doc.add_table(rows=4, cols=2)
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hdr.autofit = False

    # Row 0
    c00 = table_hdr.cell(0, 0)
    p = c00.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CÔNG TY ĐIỆN LỰC VŨNG TÀU")
    r.font.size = Pt(11)

    c01 = table_hdr.cell(0, 1)
    p = c01.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    r.bold = True
    r.font.size = Pt(11)

    # Row 1
    c10 = table_hdr.cell(1, 0)
    p = c10.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("PHÒNG TÀI CHÍNH KẾ TOÁN")
    r.bold = True
    r.font.size = Pt(11)

    c11 = table_hdr.cell(1, 1)
    p = c11.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Độc lập – Tự do – Hạnh phúc")
    r.bold = True
    r.font.size = Pt(11)

    # Row 2: đường kẻ gạch
    c20 = table_hdr.cell(2, 0)
    p = c20.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("-----------------------").font.size = Pt(10)

    c21 = table_hdr.cell(2, 1)
    p = c21.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("-----------------------").font.size = Pt(10)

    # Row 3: Số hiệu & Địa danh ngày tháng
    now_year = datetime.date.today().year
    c30 = table_hdr.cell(3, 0)
    p = c30.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Số:       /TTr-TCKT\n").font.size = Pt(11)
    r_vv = p.add_run("V/v đề nghị duyệt quyết toán danh mục SCL hoàn thành")
    r_vv.italic = True
    r_vv.font.size = Pt(10)

    c31 = table_hdr.cell(3, 1)
    p = c31.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Vũng Tàu, ngày     tháng      năm  {now_year}").italic = True

    # Kính gửi
    p_kg = doc.add_paragraph()
    p_kg.paragraph_format.space_before = Pt(20)
    r_kg = p_kg.add_run("Kính gửi : Giám đốc Công ty Điện lực Vũng Tàu")
    r_kg.bold = True

    # Căn cứ
    p_cc = doc.add_paragraph()
    p_cc.paragraph_format.space_before = Pt(10)
    p_cc.paragraph_format.line_spacing = 1.3
    p_cc.add_run("Căn cứ Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Tổng công ty Điện lực TP.HCM về việc ban hành quy định thực hiện công tác sửa chữa lớn tài sản trong Tổng công ty Điện lực Thành phố Hồ Chí Minh;\n")
    p_cc.add_run(f"Căn cứ  quyết định số {so_dt} ngày {ngay_dt} về việc phê duyệt bổ sung danh mục công trình và điều hòa kế hoạch vốn sửa chữa lớn năm {now_year} của Công ty Điện lực Vũng Tàu;\n")
    p_cc.add_run(f"Căn cứ báo cáo thẩm tra quyết toán công trình {ten}. Kính đề nghị Ông Giám đốc phê duyệt quyết toán công trình {ten}, mã công trình {ma} với tổng giá trị công trình là: ")
    r_v = p_cc.add_run(f"{_fmt_money_dot(tong_qt)} đồng.\n")
    r_v.bold = True
    p_cc.add_run(f"(Bằng chữ:  {bang_chu}).\n")
    p_cc.add_run("Trong đó:\n")

    # Bảng kê số liệu trong đó (tab indented)
    cost_lines = [
        ("Chi phí Sửa chữa", _fmt_money_dot(sc_qt), False, "đồng."),
        ("Chi phí vật tư, thiết bị", _fmt_money_dot(tb_qt), False, "đồng."),
        ("Chi phí KTCB khác", _fmt_money_dot(khac_qt), False, "đồng."),
        ("Tổng giá trị trước thuế", _fmt_money_dot(truoc_thue), True, "đồng."),
        ("Thuế GTGT", _fmt_money_dot(thue_qt), False, "đồng."),
        ("Tổng giá trị Quyết toán", _fmt_money_dot(tong_qt), True, "đồng."),
        ("Chi phí vật tư thu hồi", _fmt_money_dot(vtth_qt), False, "đồng."),
        ("Chi phí sửa chữa lớn", _fmt_money_dot(scl_qt), True, "đồng"),
    ]
    add_aligned_cost_lines_docx(doc, cost_lines, left_pt=49.65, colon_pt=205.55, right_pt=382.75)

    p_tt = doc.add_paragraph("Trân trọng./.")
    p_tt.paragraph_format.space_before = Pt(10)

    # Chân trang ký tên
    table_sig = doc.add_table(rows=1, cols=2)
    table_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_nn = table_sig.cell(0, 0)
    c_nn_p = c_nn.paragraphs[0]
    c_nn_p.add_run("Nơi nhận:\n").bold = True
    c_nn_p.add_run("- Như trên\n- Lưu: VT, TCKT. HMT.").italic = True

    c_tp = table_sig.cell(0, 1)
    c_tp_p = c_tp.paragraphs[0]
    c_tp_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tp1 = c_tp_p.add_run("TRƯỞNG PHÒNG TÀI CHÍNH KẾ TOÁN\n\n\n\n\n")
    r_tp1.bold = True
    r_tp2 = c_tp_p.add_run("Hà Thị Mai Hiên")
    r_tp2.bold = True

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# 2. BẢN THUYẾT MINH QUYẾT TOÁN
# (Khớp y chang file Tham khao/Thuyết minh quyết toán.pdf)
# ============================================================
def export_tmqt_word(main_row, ct_data, noi_dung_scl=''):
    doc = DocxDocument()
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(0.8)

    ten = str(main_row.get('Tên Công trình', ''))
    ma = str(main_row.get('Mã CT', '')) if pd.notna(main_row.get('Mã CT')) else ''
    so_dt = str(main_row.get('Số Dự toán', '')) if pd.notna(main_row.get('Số Dự toán')) else ''
    ngay_dt = _format_date_vn(main_row.get('Ngày Dự toán'))
    don_vi = str(main_row.get('Đơn vị QL', '')) if pd.notna(main_row.get('Đơn vị QL')) else 'Công ty Cổ phần Sửa chữa ô tô Tiến Phát'
    ghi_chu = str(main_row.get('Ghi chú', '')) if pd.notna(main_row.get('Ghi chú')) else 'Thuê ngoài'
    so_hd = str(main_row.get('Số Hợp đồng xây lắp', '')) if pd.notna(main_row.get('Số Hợp đồng xây lắp')) else ''
    ngay_hd = _format_date_vn(main_row.get('Ngày Hợp đồng xây lắp'))
    kc = _format_date_vn(main_row.get('Ngày khởi công'))
    ht = _format_date_vn(main_row.get('Ngày hoàn thành'))
    if '....' in kc: kc = '15/05/2026'
    if '....' in ht: ht = '30/06/2026'

    bd = get_cost_breakdown(ct_data)
    scl_dt = bd.get('SCL', {}).get('dt', 0)
    scl_qt = bd.get('SCL', {}).get('qt', 0)
    if scl_dt == 0 and 'Giá trị Dự toán' in main_row:
        scl_dt = float(main_row.get('Giá trị Dự toán', 0))
    if scl_qt == 0 and 'Giá trị Q.định phê duyệt QT công trình' in main_row:
        scl_qt = float(main_row.get('Giá trị Q.định phê duyệt QT công trình', 0))

    kh_val = main_row.get('Kế hoạch', 0)
    kh_tien = _safe_int(kh_val * 1000000) if kh_val < 100000 else _safe_int(kh_val)
    if kh_tien == 0: kh_tien = _safe_int(scl_dt)

    klcv = str(main_row.get('Khối lượng công việc', '')) if pd.notna(main_row.get('Khối lượng công việc')) else ''

    # Tiêu ngữ
    table_hdr = doc.add_table(rows=1, cols=2)
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hdr.autofit = False
    c_l = table_hdr.cell(0, 0)
    c_r = table_hdr.cell(0, 1)

    p_l = c_l.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_l.add_run("TỔNG CÔNG TY\nĐIỆN LỰC TP HỒ CHÍ MINH\n").font.size = Pt(11)
    r_l2 = p_l.add_run("CÔNG TY ĐIỆN LỰC VŨNG TÀU\n")
    r_l2.bold = True
    r_l2.font.size = Pt(11)
    p_l.add_run("Số:          /BB-PCVT").font.size = Pt(11)

    p_r = c_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_r.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n").bold = True
    p_r.add_run("Độc lập - Tự do - Hạnh Phúc\n").bold = True
    p_r.add_run("-----------------------\n").font.size = Pt(10)
    p_r.add_run(f"Vũng Tàu, ngày      tháng      năm {datetime.date.today().year}").italic = True

    # Tiêu đề
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(15)
    p_title.paragraph_format.space_after = Pt(15)
    r_t = p_title.add_run("BẢN THUYẾT MINH QUYẾT TOÁN")
    r_t.bold = True
    r_t.font.size = Pt(14)

    # Thông tin chung
    doc.add_paragraph(f"- Tên gói thầu: Cung cấp vật tư phụ tùng, thi công sửa chữa và mua bảo hiểm")
    doc.add_paragraph(f"- Công trình: “{ten}”")
    doc.add_paragraph(f"- Giá trị vốn kế hoạch: {_fmt_money_dot(kh_tien)} đồng")
    doc.add_paragraph(f"- Hình thức: {ghi_chu or 'Thuê ngoài'}")
    doc.add_paragraph(f"- Tên đơn vị thi công: {don_vi}")
    doc.add_paragraph(f"- Giá trị dự toán được duyệt (sau thuế): {_fmt_money_dot(scl_dt)} đồng")
    doc.add_paragraph(f"- Thời gian khởi công: {kc}")
    doc.add_paragraph(f"- Thời gian hoàn thành: {ht}")
    doc.add_paragraph(f"- Giá trị quyết toán danh mục hoàn thành: {_fmt_money_dot(scl_qt)} đồng")
    
    # Khối lượng
    p_kl = doc.add_paragraph()
    p_kl.add_run(f"- Khối lượng công việc chủ yếu đã hoàn thành thay thế sửa chữa, cụ thể:\n")
    if klcv:
        p_kl.add_run(klcv)
    else:
        p_kl.add_run("(Thực hiện đúng theo khối lượng biên bản nghiệm thu hoàn thành công trình)")

    # Căn cứ
    doc.add_paragraph("\n- Các căn cứ về chế độ để lập quyết toán:")
    doc.add_paragraph("  + Căn cứ Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Tổng công ty Điện lực TP.HCM về việc ban hành quy định thực hiện công tác sửa chữa lớn tài sản trong Tổng công ty Điện lực Thành phố Hồ Chí Minh.")
    if so_hd:
        doc.add_paragraph(f"  + Hợp đồng số: {so_hd} ngày {ngay_hd} giữa Công ty Điện lực Vũng Tàu và {don_vi}.")
    doc.add_paragraph(f"  + Bảng kê tổng hợp quyết toán do {don_vi} lập và được Công ty Điện lực Vũng Tàu thỏa hiệp.")

    # Phân tích & Đánh giá
    doc.add_paragraph("- Phân tích các nhân tố tăng giảm so với dự toán được duyệt: Không có")
    doc.add_paragraph("- Đánh giá hiệu quả của công việc SCL: Đảm bảo an toàn trong vận hành.")
    doc.add_paragraph("- Các kiến nghị: Không có ./.")

    # Ký tên
    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(25)
    table_sig = doc.add_table(rows=2, cols=2)
    table_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_nn = table_sig.cell(0, 0)
    c_nn.paragraphs[0].add_run("Nơi nhận:\n- P/Đ liên quan (để thực hiện);\n- Lưu: VT, TCKT, HMT.").italic = True

    c_gd = table_sig.cell(0, 1)
    c_gd_p = c_gd.paragraphs[0]
    c_gd_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_gd_p.add_run("GIÁM ĐỐC\n\n\n\n\nNguyễn Ngọc Tuyến").bold = True

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# 3. PHIẾU THẨM TRA QUYẾT TOÁN
# (Mẫu Phụ lục 10.4 - QĐ 202/QĐ-HĐTV)
# ============================================================
def export_phieu_tham_tra_word(main_row):
    doc = DocxDocument()
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(0.8)

    ten = str(main_row.get('Tên Công trình', ''))
    ma = str(main_row.get('Mã CT', '')) if pd.notna(main_row.get('Mã CT')) else ''
    don_vi = str(main_row.get('Đơn vị QL', '')) if pd.notna(main_row.get('Đơn vị QL')) else 'Công ty Điện lực Vũng Tàu'
    ghi_chu = str(main_row.get('Ghi chú', '')) if pd.notna(main_row.get('Ghi chú')) else ''
    so_hd = str(main_row.get('Số Hợp đồng xây lắp', '')) if pd.notna(main_row.get('Số Hợp đồng xây lắp')) else ''
    ngay_hd = _format_date_vn(main_row.get('Ngày Hợp đồng xây lắp'))
    so_dt = str(main_row.get('Số Dự toán', '')) if pd.notna(main_row.get('Số Dự toán')) else ''
    ngay_dt = _format_date_vn(main_row.get('Ngày Dự toán'))
    is_tu_lam = 'tự' in ghi_chu.lower() if ghi_chu else False

    table_hdr = doc.add_table(rows=1, cols=2)
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hdr.autofit = False
    c_l = table_hdr.cell(0, 0)
    c_r = table_hdr.cell(0, 1)

    p_l = c_l.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_l.add_run("TỔNG CÔNG TY\nĐIỆN LỰC TP HỒ CHÍ MINH\n").font.size = Pt(11)
    r_l2 = p_l.add_run("CÔNG TY ĐIỆN LỰC VŨNG TÀU\n")
    r_l2.bold = True
    p_l.add_run("Số:       /PTT-PCVT").font.size = Pt(11)

    p_r = c_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_r.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n").bold = True
    p_r.add_run("Độc lập - Tự do - Hạnh phúc\n").bold = True
    p_r.add_run("-----------------------\n").font.size = Pt(10)
    p_r.add_run(f"Vũng Tàu, ngày     tháng     năm {datetime.date.today().year}").italic = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(15)
    r_t = p_title.add_run("PHIẾU THẨM TRA QUYẾT TOÁN\nCÔNG TRÌNH SỬA CHỮA LỚN")
    r_t.bold = True
    r_t.font.size = Pt(14)

    doc.add_paragraph(f"1. Tên công trình SCL: {ten}")
    doc.add_paragraph(f"2. Mã công trình: {ma}")
    doc.add_paragraph(f"3. Đơn vị quản lý: {don_vi}")
    if is_tu_lam:
        doc.add_paragraph(f"4. Phương thức chọn thầu thực hiện: Tự làm")
    else:
        doc.add_paragraph(f"4. Phương thức chọn thầu thực hiện: Thuê ngoài (Hợp đồng số {so_hd} ngày {ngay_hd})")
    doc.add_paragraph(f"5. Đơn vị thực hiện: {don_vi}")
    doc.add_paragraph(f"6. Căn cứ phê duyệt PAKT-DT: Quyết định số {so_dt} ngày {ngay_dt} của Công ty Điện lực Vũng Tàu")

    doc.add_paragraph("\n7. Ý KIẾN THẨM TRA CỦA CÁC ĐƠN VỊ THÀNH VIÊN:")
    doc.add_paragraph("- Phòng Kỹ thuật: Đã kiểm tra khối lượng hoàn công thực tế phù hợp với PAKT-DT được duyệt.")
    doc.add_paragraph("- Phòng Kế hoạch Vật tư: Đã rà soát chi phí vật tư, hợp đồng và đối chiếu VTTB thu hồi.")
    doc.add_paragraph("- Phòng Tài chính Kế toán: Đã kiểm tra tính hợp pháp của chứng từ, hóa đơn GTGT và đối chiếu sổ sách kế toán.")

    doc.add_paragraph("\n8. KẾT LUẬN VÀ KIẾN NGHỊ:")
    doc.add_paragraph("Hồ sơ quyết toán công trình SCL đầy đủ tính pháp lý, tuân thủ đúng quy trình quản lý SCL theo Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của EVNHCMC. Kính trình Giám đốc Công ty phê duyệt.")

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(20)
    table_sig = doc.add_table(rows=2, cols=3)
    table_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    cols_hdr = ["ĐẠI DIỆN P.KHVT", "ĐẠI DIỆN P.TCKT", "TỔ TRƯỞNG TỔ THẨM TRA"]
    for i, h in enumerate(cols_hdr):
        c = table_sig.cell(0, i)
        cp = c.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.add_run(h).bold = True
        c2 = table_sig.cell(1, i)
        c2.paragraphs[0].paragraph_format.space_before = Pt(35)

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# 4. BÁO CÁO THẨM TRA QUYẾT TOÁN CÔNG TRÌNH SCL
# (Khớp y chang file HS_mau/Báo cáo Thẩm tra.docx)
# ============================================================
def export_bao_cao_tham_tra_word(main_row, ct_data, thqt_data=None):
    doc = DocxDocument()
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(0.8)

    ten = str(main_row.get('Tên Công trình', '')) if main_row is not None else ''
    ma = str(main_row.get('Mã CT', '')) if main_row is not None and pd.notna(main_row.get('Mã CT')) else ''
    don_vi_th = str(main_row.get('Đơn vị QL', '')) if main_row is not None and pd.notna(main_row.get('Đơn vị QL')) else 'Công ty CP Tư vấn Dịch vụ Xây dựng Thái Bình'
    ghi_chu = str(main_row.get('Ghi chú', '')) if main_row is not None and pd.notna(main_row.get('Ghi chú')) else 'đấu thầu rộng rãi'
    so_hd = str(main_row.get('Số Hợp đồng xây lắp', '')) if main_row is not None and pd.notna(main_row.get('Số Hợp đồng xây lắp')) else '40-2026/HĐXL/PCVT-TB'

    if thqt_data and 'summary' in thqt_data:
        sm = thqt_data['summary']
        if thqt_data.get('ten_ct'): ten = thqt_data['ten_ct']
        if thqt_data.get('ma_ct'): ma = thqt_data['ma_ct']
        sc_dt = sm.get('sc_dt', 0)
        sc_qt = sm.get('sc_qt', 0)
        tb_dt = sm.get('tb_dt', 0)
        tb_qt = sm.get('tb_qt', 0)
        khac_dt = sm.get('khac_dt', 0)
        khac_qt = sm.get('khac_qt', 0)
        truoc_thue = sm.get('truoc_thue_qt', 0)
        thue_qt = sm.get('thue_qt', 0)
        tong_qt = sm.get('tong_qt', 0)
        vtth_qt = sm.get('vtth_qt', 0)
        scl_qt = sm.get('scl_qt', 0)
        ktcb_list = thqt_data.get('ktcb_rows', [])
        tot_ktcb_dt = sm.get('tot_ktcb_dt', 0)
        tot_ktcb_qt = sm.get('tot_ktcb_qt', 0)
    else:
        bd = get_cost_breakdown(ct_data)
        scl_dt = bd.get('SCL', {}).get('dt', 0)
        scl_qt = bd.get('SCL', {}).get('qt', 0)
        if scl_dt == 0 and main_row is not None and 'Giá trị Dự toán' in main_row:
            scl_dt = float(main_row.get('Giá trị Dự toán', 0))
        if scl_qt == 0 and main_row is not None and 'Giá trị Q.định phê duyệt QT công trình' in main_row:
            scl_qt = float(main_row.get('Giá trị Q.định phê duyệt QT công trình', 0))

        sc_dt = bd.get('B', {}).get('dt', 0)
        sc_qt = bd.get('B', {}).get('qt', 0)
        tb_dt = bd.get('A', {}).get('dt', 0)
        tb_qt = bd.get('A', {}).get('qt', 0)
        khac_dt = bd.get('C', {}).get('dt', 0)
        khac_qt = bd.get('C', {}).get('qt', 0)
        vtth_qt = bd.get('F', {}).get('qt', 0)
        thue_qt = bd.get('B.8', {}).get('qt', 0)
        if thue_qt == 0 and scl_qt > 0: thue_qt = int(round(scl_qt - scl_qt / 1.08))
        truoc_thue = scl_qt - thue_qt if scl_qt >= thue_qt else scl_qt
        tong_qt = scl_qt
        ktcb_list = []
        tot_ktcb_dt = khac_dt
        tot_ktcb_qt = khac_qt

    now_year = datetime.date.today().year

    # Header 2 cột
    table_hdr = doc.add_table(rows=1, cols=2)
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hdr.autofit = False
    c_l = table_hdr.cell(0, 0)
    c_r = table_hdr.cell(0, 1)

    p_l = c_l.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_l.add_run("TỔNG CÔNG TY ĐIỆN LỰC TP HỒ CHÍ MINH\n").font.size = Pt(10)
    r_l2 = p_l.add_run("CÔNG TY ĐIỆN LỰC VŨNG TÀU\n")
    r_l2.bold = True
    r_l2.font.size = Pt(11)
    p_l.add_run("Số:       /BC-TCKT").font.size = Pt(11)

    p_r = c_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_r.add_run("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM\n").bold = True
    p_r.add_run("Độc lập - Tự do - Hạnh phúc\n").bold = True
    p_r.add_run("-----------------------\n").font.size = Pt(10)
    p_r.add_run(f"Vũng Tàu, ngày       tháng      năm {now_year}.").italic = True

    # Tiêu đề
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(15)
    p_title.paragraph_format.space_after = Pt(12)
    r_t = p_title.add_run("BÁO CÁO THẨM TRA QUYẾT TOÁN\nCÔNG TRÌNH SỬA CHỮA LỚN")
    r_t.bold = True
    r_t.font.size = Pt(14)

    # Căn cứ
    p_cc = doc.add_paragraph()
    p_cc.paragraph_format.line_spacing = 1.3
    p_cc.add_run(f"Căn cứ hồ sơ quyết toán công trình: {ten}.\n")
    p_cc.add_run("Tổ thẩm tra quyết toán công trình sửa chữa lớn Công ty Điện lực Vũng Tàu thẩm tra hồ sơ quyết toán công trình với kết quả như sau:")

    # I/- Nội dung thẩm tra
    p_i = doc.add_paragraph()
    p_i.paragraph_format.space_before = Pt(6)
    r_i = p_i.add_run("I/- Nội dung thẩm tra:")
    r_i.bold = True
    r_i.italic = True

    p_nd = doc.add_paragraph()
    p_nd.paragraph_format.left_indent = Inches(0.25)
    p_nd.paragraph_format.line_spacing = 1.3
    p_nd.add_run(f"Tên công trình: {ten}.\n")
    p_nd.add_run(f"Đơn vị quản lý: Công ty Điện lực Vũng Tàu.\n")
    p_nd.add_run(f"Phương thức chọn thầu thực hiện: {ghi_chu}\n")
    p_nd.add_run("Hợp đồng số:\n")
    p_nd.add_run(f"{so_hd}.\n")
    p_nd.add_run("Đơn vị thực hiện:\n")
    p_nd.add_run(f"{don_vi_th}.")

    # Nội dung cụ thể
    p_ct = doc.add_paragraph()
    p_ct.paragraph_format.space_before = Pt(6)
    p_ct.add_run("Nội dung cụ thể:").bold = True

    # 1/ Phần sửa chữa
    p_sc = doc.add_paragraph()
    p_sc.paragraph_format.left_indent = Inches(0.2)
    p_sc.paragraph_format.space_after = Pt(2)
    p_sc.add_run("1/ Phần sửa chữa:").bold = True

    sc_items = [
        ("- Số dự toán", _fmt_money_dot(sc_dt), False, "đồng."),
        ("- Số quyết toán", _fmt_money_dot(sc_qt), False, "đồng."),
        ("- Số thẩm tra", _fmt_money_dot(sc_qt), False, "đồng."),
        ("- Số chênh lệch TTr/QT", "0", False, "đồng."),
    ]
    add_aligned_cost_lines_docx(doc, sc_items, left_pt=35.45, colon_pt=177.2, right_pt=340.2)

    # 2/ Phần thiết bị
    p_tb = doc.add_paragraph()
    p_tb.paragraph_format.left_indent = Inches(0.2)
    p_tb.paragraph_format.space_after = Pt(2)
    p_tb.add_run("2/  Phần vật tư, thiết bị:").bold = True

    tb_items = [
        ("- Số dự toán", _fmt_money_dot(tb_dt), False, "đồng."),
        ("- Số quyết toán", _fmt_money_dot(tb_qt), False, "đồng."),
        ("- Số thẩm tra", _fmt_money_dot(tb_qt), False, "đồng."),
        ("- Số chênh lệch TTr/QT", "0", False, "đồng."),
    ]
    add_aligned_cost_lines_docx(doc, tb_items, left_pt=35.45, colon_pt=177.2, right_pt=340.2)

    # 3/ Phần kiến thiết cơ bản khác
    p_kc = doc.add_paragraph()
    p_kc.paragraph_format.space_before = Pt(6)
    p_kc.add_run("3/ Phần kiến thiết cơ bản khác:").bold = True

    # Bảng chi phí KTCB khác (5 cột chuẩn xác)
    t_ktcb = doc.add_table(rows=1, cols=5)
    t_ktcb.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_ktcb.autofit = False
    hdr_kt = t_ktcb.rows[0].cells
    hdr_names = ["Nội dung", "Dự toán được duyệt", "Giá trị quyết toán", "Số thẩm tra", "Chênh lệch TTr/QT"]
    for i, hn in enumerate(hdr_names):
        hdr_kt[i].text = hn
        hdr_kt[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_kt[i].paragraphs[0].runs[0].bold = True

    # Default items if not provided by thqt
    if not ktcb_list:
        ktcb_list = [
            {"name": "Chi phí Lập khảo sát, Lập PAKT-DT", "dt": 46296296, "qt": 46296296, "tt": 46296296, "diff": 0},
            {"name": "2- Chi phí giám sát công trình", "dt": 38888889, "qt": 38888889, "tt": 38888889, "diff": 0},
            {"name": "Chi phí bảo hiểm công trình", "dt": 898673, "qt": 0, "tt": 0, "diff": 0},
            {"name": "Chi phí thẩm tra thiết kế", "dt": 2107797, "qt": 0, "tt": 0, "diff": 0},
            {"name": "Chi phí thẩm tra dự toán", "dt": 2042439, "qt": 0, "tt": 0, "diff": 0},
            {"name": "Chi phí thẩm tra phê duyệt quyết toán", "dt": 5619375, "qt": 0, "tt": 0, "diff": 0},
            {"name": "Chi phí dự phòng", "dt": 49292765, "qt": 0, "tt": 0, "diff": 0},
        ]
        tot_ktcb_dt = sum(x["dt"] for x in ktcb_list)
        tot_ktcb_qt = sum(x["qt"] for x in ktcb_list)

    for item in ktcb_list:
        r_c = t_ktcb.add_row().cells
        r_c[0].text = item["name"]
        r_c[1].text = _fmt_money_dot(item.get("dt", 0))
        r_c[2].text = _fmt_money_dot(item.get("qt", 0))
        r_c[3].text = _fmt_money_dot(item.get("tt", item.get("qt", 0)))
        r_c[4].text = "0"
        for cell in r_c[1:]:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Row tổng cộng
    r_tot = t_ktcb.add_row().cells
    r_tot[0].text = "Tổng cộng"
    r_tot[0].paragraphs[0].runs[0].bold = True
    r_tot[1].text = _fmt_money_dot(tot_ktcb_dt)
    r_tot[2].text = _fmt_money_dot(tot_ktcb_qt)
    r_tot[3].text = _fmt_money_dot(tot_ktcb_qt)
    r_tot[4].text = "0"
    for cell in r_tot[1:]:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True

    # II/ Kết luận
    p_ii = doc.add_paragraph()
    p_ii.paragraph_format.space_before = Pt(12)
    r_ii = p_ii.add_run("II/ Kết luận:")
    r_ii.bold = True
    r_ii.italic = True

    p_kl = doc.add_paragraph()
    p_kl.paragraph_format.left_indent = Inches(0.2)
    p_kl.add_run(f"Sau khi xem xét thẩm tra hồ sơ, tổ thẩm tra quyết toán chấp thuận tổng giá trị quyết toán công trình nêu trên là: ")
    p_kl.add_run(f"{_fmt_money_dot(tong_qt)} đồng").bold = True
    p_kl.add_run(", cụ thể:\n")

    summary_items = [
        ("Chi phí Sửa chữa", _fmt_money_dot(sc_qt), False, "đồng."),
        ("Chi phí vật tư, thiết bị", _fmt_money_dot(tb_qt), False, "đồng."),
        ("Chi phí KTCB khác", _fmt_money_dot(khac_qt), False, "đồng."),
        ("Tổng giá trị trước thuế", _fmt_money_dot(truoc_thue), True, "đồng."),
        ("Thuế GTGT", _fmt_money_dot(thue_qt), False, "đồng."),
        ("Tổng giá trị Quyết toán", _fmt_money_dot(tong_qt), True, "đồng."),
        ("Chi phí vật tư thu hồi", _fmt_money_dot(vtth_qt), False, "đồng."),
        ("Chi phí sửa chữa lớn", _fmt_money_dot(scl_qt), True, "đồng"),
    ]
    add_aligned_cost_lines_docx(doc, summary_items, left_pt=49.65, colon_pt=205.55, right_pt=382.75)

    p_propose = doc.add_paragraph("Kính trình và đề nghị Ông Giám đốc phê duyệt.")
    p_propose.paragraph_format.space_before = Pt(8)

    # Danh sách thành viên tổ thẩm tra
    p_tv = doc.add_paragraph()
    p_tv.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tv.paragraph_format.space_before = Pt(14)
    p_tv.paragraph_format.space_after = Pt(10)
    p_tv.add_run("THÀNH VIÊN TỔ THẨM TRA QUYẾT TOÁN").bold = True

    # Bảng 8 rows x 2 cols như mẫu Word
    table_tv = doc.add_table(rows=8, cols=2)
    table_tv.alignment = WD_TABLE_ALIGNMENT.CENTER
    members = [
        "Hà Thị Mai Hiên (KTT)…………………….…...........",
        "Nguyễn văn Quyến (TP.QLĐT)………………………",
        "Nguyễn Mạnh Hiệp (TP.KHVT)……………………...",
        "Nguyễn Đức Minh (TP.KTAT)……………………….",
        "Hoàng Minh Tuấn (CV.TCKT)…………….…………",
        "Nguyễn Văn Huy (TT.KTAT)………………………",
        "Lê Anh Ngọc (TT.KHVT)…………………………….",
        "Lê Hoài Nam (CV.QLĐT)………………………",
    ]
    for idx, mem in enumerate(members):
        c_left = table_tv.cell(idx, 0)
        p = c_left.paragraphs[0]
        p.add_run(mem)
        p.paragraph_format.space_after = Pt(2)

    # Cột bên phải: TỔ TRƯỞNG TỔ THẨM TRA ở row 0 và Trần Thanh Hải ở row 4
    c_r0 = table_tv.cell(0, 1)
    p_r0 = c_r0.paragraphs[0]
    p_r0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tt = p_r0.add_run("TỔ TRƯỞNG TỔ THẨM TRA")
    r_tt.bold = True

    c_r4 = table_tv.cell(4, 1)
    p_r4 = c_r4.paragraphs[0]
    p_r4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = p_r4.add_run("Trần Thanh Hải")
    r_name.bold = True

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# 5. QUYẾT ĐỊNH PHÊ DUYỆT QUYẾT TOÁN CÔNG TRÌNH SCL HOÀN THÀNH
# (Khớp y chang file HS_mau/Quyết định phê duyệt quyết toán CT SCL-VTAD2607001.doc)
# ============================================================
def export_qd_phe_duyet_word(main_row, ct_data, thqt_data=None):
    doc = DocxDocument()
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(0.8)

    ten = str(main_row.get('Tên Công trình', '')) if main_row is not None else ''
    ma = str(main_row.get('Mã CT', '')) if main_row is not None and pd.notna(main_row.get('Mã CT')) else ''
    so_dt = str(main_row.get('Số Dự toán', '')) if main_row is not None and pd.notna(main_row.get('Số Dự toán')) else '135/QĐ-PCVT'
    ngay_dt = _format_date_vn(main_row.get('Ngày Dự toán')) if main_row is not None else '16/03/2026'
    if '....' in ngay_dt: ngay_dt = '16/03/2026'

    if thqt_data and 'summary' in thqt_data:
        sm = thqt_data['summary']
        if thqt_data.get('ten_ct'): ten = thqt_data['ten_ct']
        if thqt_data.get('ma_ct'): ma = thqt_data['ma_ct']
        sc_qt = sm.get('sc_qt', 0)
        tb_qt = sm.get('tb_qt', 0)
        khac_qt = sm.get('khac_qt', 0)
        truoc_thue = sm.get('truoc_thue_qt', 0)
        thue_qt = sm.get('thue_qt', 0)
        tong_qt = sm.get('tong_qt', 0)
        vtth_qt = sm.get('vtth_qt', 0)
        scl_qt = sm.get('scl_qt', 0)
        bang_chu = sm.get('bang_chu') or doc_so_vn(tong_qt)
    else:
        bd = get_cost_breakdown(ct_data)
        scl_qt = bd.get('SCL', {}).get('qt', 0)
        if scl_qt == 0 and main_row is not None and 'Giá trị Q.định phê duyệt QT công trình' in main_row:
            scl_qt = float(main_row.get('Giá trị Q.định phê duyệt QT công trình', 0))

        tb_qt = bd.get('A', {}).get('qt', 0)
        sc_qt = bd.get('B', {}).get('qt', 0)
        khac_qt = bd.get('C', {}).get('qt', 0)
        vtth_qt = bd.get('F', {}).get('qt', 0)
        thue_qt = bd.get('B.8', {}).get('qt', 0)

        if thue_qt == 0 and scl_qt > 0: thue_qt = int(round(scl_qt - scl_qt / 1.08))
        truoc_thue = scl_qt - thue_qt if scl_qt >= thue_qt else scl_qt
        tong_qt = scl_qt
        bang_chu = doc_so_vn(scl_qt)

    now_year = datetime.date.today().year

    # Tiêu ngữ (5 rows x 2 cols giống HS_mau)
    table_hdr = doc.add_table(rows=5, cols=2)
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hdr.autofit = False

    # Row 0
    c00 = table_hdr.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p00.add_run("TỔNG CÔNG TY").font.size = Pt(11)

    c01 = table_hdr.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ch = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    r_ch.bold = True
    r_ch.font.size = Pt(11)

    # Row 1
    c10 = table_hdr.cell(1, 0)
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p10.add_run("ĐIỆN LỰC TP HỒ CHÍ MINH").font.size = Pt(11)

    c11 = table_hdr.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_dl = p11.add_run("Độc lập – Tự do – Hạnh phúc")
    r_dl.bold = True
    r_dl.font.size = Pt(11)

    # Row 2
    c20 = table_hdr.cell(2, 0)
    p20 = c20.paragraphs[0]
    p20.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_pc = p20.add_run("CÔNG TY ĐIỆN LỰC VŨNG TÀU")
    r_pc.bold = True
    r_pc.font.size = Pt(11)

    # Row 3 (line separator)
    c31 = table_hdr.cell(3, 1)
    p31 = c31.paragraphs[0]
    p31.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p31.add_run("-----------------------").font.size = Pt(10)

    # Row 4: Số hiệu & Địa danh ngày tháng
    c40 = table_hdr.cell(4, 0)
    p40 = c40.paragraphs[0]
    p40.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p40.add_run("Số:       /QĐ-PCVT").font.size = Pt(11)

    c41 = table_hdr.cell(4, 1)
    p41 = c41.paragraphs[0]
    p41.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p41.add_run(f"Vũng Tàu, ngày     tháng      năm  {now_year}").italic = True

    # Tiêu đề
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(15)
    r_t = p_title.add_run("QUYẾT ĐỊNH\n")
    r_t.bold = True
    r_t.font.size = Pt(14)
    r_sub = p_title.add_run("V/v phê duyệt quyết toán công trình Sửa chữa lớn hoàn thành\n")
    r_sub.font.size = Pt(12)

    # Thông tin đầu trang
    doc.add_paragraph(f"Tên công trình: {ten}")
    doc.add_paragraph(f"Mã công trình: {ma}")
    doc.add_paragraph("Chủ đầu tư: Công ty Điện lực Vũng Tàu")
    doc.add_paragraph(f"Kế hoạch vốn năm: {now_year}")
    doc.add_paragraph("Nguồn vốn: Sửa chữa lớn")

    # Giám đốc & Căn cứ
    p_gd = doc.add_paragraph()
    p_gd.paragraph_format.space_before = Pt(10)
    p_gd.add_run("GIÁM ĐỐC CÔNG TY ĐIỆN LỰC VŨNG TÀU\n").bold = True
    p_gd.paragraph_format.line_spacing = 1.3
    p_gd.add_run("Căn cứ Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Tổng công ty Điện lực TP.HCM về việc ban hành quy định thực hiện công tác sửa chữa lớn tài sản trong Tổng công ty Điện lực Thành phố Hồ Chí Minh;\n")
    p_gd.add_run(f"Căn cứ  quyết định số {so_dt} ngày {ngay_dt} về việc phê duyệt bổ sung danh mục công trình và điều hòa kế hoạch vốn sửa chữa lớn năm {now_year} của Công ty Điện lực Vũng Tàu;\n")
    p_gd.add_run(f"Căn cứ hồ sơ quyết toán công trình: {ten};\n")
    p_gd.add_run(f"Căn cứ báo cáo thẩm tra quyết toán của Tổ thẩm tra duyệt quyết toán công trình sửa chữa lớn Công ty Điện lực Vũng tàu.")

    p_qd = doc.add_paragraph()
    p_qd.paragraph_format.space_before = Pt(10)
    p_qd.add_run("QUYẾT ĐỊNH :\n").bold = True

    # Điều 1
    p_d1 = doc.add_paragraph()
    p_d1.add_run("Điều 1. Phê duyệt quyết toán\n").bold = True
    p_d1.add_run(f"Công trình: {ten} với tổng giá trị công trình: ")
    p_d1.add_run(f"{_fmt_money_dot(tong_qt)} đồng.\n").bold = True
    p_d1.add_run(f"(Bằng chữ: {bang_chu}).\n")
    p_d1.add_run("Trong đó:\n")

    summary_items = [
        ("Chi phí Sửa chữa", _fmt_money_dot(sc_qt), False, "đồng."),
        ("Chi phí vật tư, thiết bị", _fmt_money_dot(tb_qt), False, "đồng."),
        ("Chi phí KTCB khác", _fmt_money_dot(khac_qt), False, "đồng."),
        ("Tổng giá trị trước thuế", _fmt_money_dot(truoc_thue), True, "đồng."),
        ("Thuế GTGT", _fmt_money_dot(thue_qt), False, "đồng."),
        ("Tổng giá trị Quyết toán", _fmt_money_dot(tong_qt), True, "đồng."),
        ("Chi phí vật tư thu hồi", _fmt_money_dot(vtth_qt), False, "đồng."),
        ("Chi phí sửa chữa lớn", _fmt_money_dot(scl_qt), True, "đồng"),
    ]
    add_aligned_cost_lines_docx(doc, summary_items, left_pt=49.65, colon_pt=205.55, right_pt=382.75)

    # Điều 2
    p_d2 = doc.add_paragraph()
    p_d2.paragraph_format.space_before = Pt(8)
    p_d2.add_run("Điều 2. Nguồn vốn thực hiện công trình\n").bold = True
    p_d2.add_run(f"Nguồn vốn: Sửa chữa lớn năm {now_year} của Công ty Điện lực Vũng Tàu.")

    # Điều 3
    p_d3 = doc.add_paragraph()
    p_d3.paragraph_format.space_before = Pt(8)
    p_d3.add_run("Điều 3. Thực hiện\n").bold = True
    p_d3.add_run("Phòng Tài chính Kế toán, Phòng Kế hoạch Vật tư, Phòng Quản lý đầu tư, Phòng Kỹ Thuật An Toàn Công ty Điện lực Vũng Tàu chịu trách nhiệm thi hành quyết định này./.")

    # Ký tên
    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(15)
    table_sig = doc.add_table(rows=1, cols=2)
    table_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_nn = table_sig.cell(0, 0)
    c_nn.paragraphs[0].add_run("Nơi nhận:\n- P.KHVT, P.QLĐT, P.KTAT (để thực hiện)\n- Lưu: VT, TCKT, HMT. (04)").italic = True

    c_gd = table_sig.cell(0, 1)
    c_gd_p = c_gd.paragraphs[0]
    c_gd_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_gdt = c_gd_p.add_run("GIÁM ĐỐC\n\n\n\n\n")
    r_gdt.bold = True
    r_gdn = c_gd_p.add_run("Nguyễn Ngọc Tuyến")
    r_gdn.bold = True

    _apply_font(doc)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
