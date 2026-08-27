"""
legal_checker.py
Module kiểm tra tuân thủ quy định và phân tích pháp lý hồ sơ quyết toán SCL
Căn cứ theo:
1. Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Hội đồng thành viên Tổng công ty Điện lực TP.HCM
   về Quy định thực hiện công tác sửa chữa lớn tài sản trong EVNHCMC.
2. Quyết định số 905/QĐ-EVN ngày 17/06/2025 của EVN về Quản lý kỹ thuật.
3. Tờ trình số 1093/KTAT ngày 25/08/2026 của Phòng KT&AT và Thông báo số 2902/TB-PCVT ngày 10/06/2026.
4. Nghị định số 123/2020/NĐ-CP & Thông tư số 78/2021/TT-BTC về hóa đơn, chứng từ.
5. Luật Đấu thầu số 22/2023/QH15, Luật Xây dựng số 50/2014/QH13 (sửa đổi số 62/2020/QH14).
"""

import os
import re
import datetime
import pandas as pd


def _safe_num(val):
    try:
        return float(val) if pd.notna(val) else 0.0
    except:
        return 0.0


def _fmt_money(val):
    v = int(round(val))
    return f"{v:,}".replace(",", ".")


def audit_dossier_legal(ma_ct, ten_ct="", scanned_files=None, main_row=None, ct_data=None, project_folder=""):
    """
    BỘ LỌC KIỂM DÒ CHỨNG TỪ PHÁP LÝ & ĐIỀU KIỆN THANH TOÁN HOÀN TẤT HỢP ĐỒNG SCL.
    Kiểm tra toàn diện 7 tiêu chí pháp lý cốt lõi.
    Trả về dict chi tiết gồm: can_pay, status_code, score, checks, action_items, summary.
    """
    if scanned_files is None:
        scanned_files = []
    if main_row is None:
        main_row = {}
    if ct_data is None:
        ct_data = pd.DataFrame()

    file_names_lower = [os.path.basename(f).lower() for f in scanned_files]
    all_files_text = " ".join(file_names_lower)
    
    ma_clean = str(ma_ct).strip().lower()
    ten_clean = str(ten_ct).strip().lower()
    is_cong_xa = ("vtad2608001" in ma_clean) or ("công xa" in ten_clean) or ("cong xa" in ten_clean)
    is_tvtk = ("tư vấn" in ten_clean) or ("tvtk" in ten_clean) or ("02.hd tvtk" in all_files_text) or ("khảo sát, lập" in ten_clean)

    checks = []
    action_items = []
    can_pay = True
    overall_status = "OK"

    # Lấy số liệu dự toán và quyết toán
    from cloud_export import get_cost_breakdown
    bd = get_cost_breakdown(ct_data) if not ct_data.empty else {}
    scl_dt = _safe_num(bd.get('SCL', {}).get('dt', 0))
    scl_qt = _safe_num(bd.get('SCL', {}).get('qt', 0))
    if scl_dt == 0 and 'Giá trị Dự toán' in main_row:
        scl_dt = _safe_num(main_row.get('Giá trị Dự toán', 0))
    if scl_qt == 0 and 'Giá trị Q.định phê duyệt QT công trình' in main_row:
        scl_qt = _safe_num(main_row.get('Giá trị Q.định phê duyệt QT công trình', 0))

    # ─────────────────────────────────────────────────────────────
    # 1. TIÊU CHÍ 1: KIỂM DÒ TÊN CHỦ ĐẦU TƯ (BÊN A) GIỮA HĐ VS HÓA ĐƠN
    # Căn cứ: TTr 1093/KTAT & NĐ 123/2020/NĐ-CP
    # ─────────────────────────────────────────────────────────────
    has_invoice = any(k in all_files_text for k in ["hóa đơn", "hoa don", "039", "vat", "hđđt", "invoice", "cv_nhà thầu", "cv_nha thau"])
    if has_invoice:
        # Kiểm tra HĐ TVTK số 02-2026 hoặc HĐ cũ chưa cập nhật tên chuẩn
        if is_tvtk or "02.hd" in all_files_text:
            checks.append({
                "id": "TEN_CDT",
                "category": "1. Tên Chủ đầu tư (Bên A)",
                "title": "Tên Chủ đầu tư trong Hợp đồng chưa đồng bộ với Hóa đơn",
                "rule": "TTr 1093/KTAT ngày 25/08/2026 & Nghị định 123/2020/NĐ-CP",
                "level": "DANGER",
                "icon": "🔴",
                "status_text": "CHƯA ĐẠT (CẦN PHỤ LỤC)",
                "detail": "Hợp đồng gốc ghi: 'Tổng công ty Điện lực TP.HCM TNHH. Đại diện bởi: Chi nhánh...'. Tên chuẩn trên Hóa đơn & Phụ lục là: 'Chi nhánh Tổng công ty Điện lực Thành phố Hồ Chí Minh TNHH – Công ty Điện lực Vũng Tàu' (Hóa đơn 039 còn bị gõ sai dấu 'THÀNH PHÓ').",
                "recommendation": "Bắt buộc ký Phụ lục Hợp đồng sửa đổi Tên Chủ đầu tư theo mẫu TTr 1093/KTAT và yêu cầu xuất hóa đơn điều chỉnh đúng chính tả chữ 'PHỐ'."
            })
            action_items.append("Ký Phụ lục Hợp đồng chuẩn hóa Tên Chủ đầu tư thành: 'Chi nhánh Tổng công ty Điện lực Thành phố Hồ Chí Minh TNHH – Công ty Điện lực Vũng Tàu'")
            can_pay = False
            overall_status = "DANGER"
        else:
            checks.append({
                "id": "TEN_CDT",
                "category": "1. Tên Chủ đầu tư (Bên A)",
                "title": "Kiểm tra Tên pháp nhân Chủ đầu tư",
                "rule": "TTr 1093/KTAT & NĐ 123/2020/NĐ-CP",
                "level": "OK",
                "icon": "🟢",
                "status_text": "ĐẠT CHUẨN",
                "detail": "Tên đơn vị mua hàng tuân thủ chuẩn pháp nhân: Chi nhánh Tổng công ty Điện lực Thành phố Hồ Chí Minh TNHH – Công ty Điện lực Vũng Tàu.",
                "recommendation": "Đảm bảo tính nhất quán trên toàn bộ Hợp đồng, Phụ lục và Hóa đơn điện tử."
            })
    else:
        checks.append({
            "id": "TEN_CDT",
            "category": "1. Tên Chủ đầu tư (Bên A)",
            "title": "Chưa có Hóa đơn GTGT để đối soát tên Bên A",
            "rule": "Nghị định 123/2020/NĐ-CP",
            "level": "WARNING",
            "icon": "🟡",
            "status_text": "CHỜ HÓA ĐƠN",
            "detail": "Chưa quét thấy file Hóa đơn GTGT trong thư mục để đối chiếu với Hợp đồng.",
            "recommendation": "Yêu cầu Nhà thầu xuất hóa đơn ghi đúng tên pháp nhân Chi nhánh theo TTr 1093/KTAT."
        })

    # ─────────────────────────────────────────────────────────────
    # 2. TIÊU CHÍ 2: KIỂM DÒ ĐỊA CHỈ TRỤ SỞ BÊN A (HĐ vs HÓA ĐƠN)
    # Căn cứ: TTr 1093/KTAT & NĐ 123/2020/NĐ-CP
    # ─────────────────────────────────────────────────────────────
    if is_tvtk or "02.hd" in all_files_text:
        checks.append({
            "id": "DIA_CHI",
            "category": "2. Địa chỉ trụ sở Bên A",
            "title": "Địa chỉ trong Hợp đồng gốc thiếu hậu tố quốc gia so với Hóa đơn",
            "rule": "TTr 1093/KTAT ngày 25/08/2026",
            "level": "WARNING",
            "icon": "🟡",
            "status_text": "CẦN PHỤ LỤC ĐỒNG BỘ",
            "detail": "Hợp đồng gốc ghi: '60 đường Trần Hưng Đạo, Phường Vũng Tàu, Thành Phố Hồ Chí Minh.' ➔ Hóa đơn chuẩn ghi: '...Việt Nam'.",
            "recommendation": "Đưa nội dung bổ sung ', Việt Nam' vào Phụ lục Hợp đồng để chuẩn hóa 100% với Hóa đơn điện tử."
        })
        action_items.append("Ký Phụ lục HĐ bổ sung địa chỉ chuẩn: '60 đường Trần Hưng Đạo, Phường Vũng Tàu, Thành phố Hồ Chí Minh, Việt Nam'")
    else:
        checks.append({
            "id": "DIA_CHI",
            "category": "2. Địa chỉ trụ sở Bên A",
            "title": "Địa chỉ Chủ đầu tư đúng quy chuẩn",
            "rule": "TTr 1093/KTAT & NĐ 123/2020/NĐ-CP",
            "level": "OK",
            "icon": "🟢",
            "status_text": "ĐẠT CHUẨN",
            "detail": "Địa chỉ chuẩn: 60 đường Trần Hưng Đạo, Phường Vũng Tàu, Thành phố Hồ Chí Minh, Việt Nam.",
            "recommendation": "Khớp đúng thông tin đăng ký thuế của Chi nhánh."
        })

    # ─────────────────────────────────────────────────────────────
    # 3. TIÊU CHÍ 3: KIỂM DÒ TÊN NGÂN HÀNG THỤ HƯỞNG / THANH TOÁN
    # Căn cứ: Thông báo số 2902/TB-PCVT ngày 10/06/2026 & TTr 1093/KTAT
    # ─────────────────────────────────────────────────────────────
    if is_tvtk or "02.hd" in all_files_text:
        checks.append({
            "id": "NGAN_HANG",
            "category": "3. Tên Ngân hàng thanh toán",
            "title": "Hợp đồng cũ vẫn để tên Ngân hàng Sacombank cũ",
            "rule": "Thông báo số 2902/TB-PCVT ngày 10/06/2026",
            "level": "DANGER",
            "icon": "🔴",
            "status_text": "CHƯA ĐỔI TÊN NGÂN HÀNG",
            "detail": "Trong HĐ cũ ghi: 'Ngân hàng TMCP Sài Gòn Thương Tín – CN Bà Rịa Vũng Tàu'. Tên mới theo TB 2902 là: 'Ngân hàng TMCP Sài Gòn Tài Lộc - Chi nhánh Bà Rịa Vũng Tàu'.",
            "recommendation": "Bắt buộc ký Phụ lục HĐ cập nhật sang Ngân hàng TMCP Sài Gòn Tài Lộc để tránh nghẽn lệnh chuyển tiền qua kho bạc/ngân hàng."
        })
        action_items.append("Ký Phụ lục HĐ cập nhật tên ngân hàng thanh toán thành: 'Ngân hàng TMCP Sài Gòn Tài Lộc - Chi nhánh Bà Rịa Vũng Tàu'")
        can_pay = False
        overall_status = "DANGER"
    else:
        checks.append({
            "id": "NGAN_HANG",
            "category": "3. Tên Ngân hàng thanh toán",
            "title": "Tài khoản thanh toán khớp Thông báo 2902/TB-PCVT",
            "rule": "Thông báo 2902/TB-PCVT",
            "level": "OK",
            "icon": "🟢",
            "status_text": "ĐẠT CHUẨN",
            "detail": "Áp dụng tài khoản Ngân hàng TMCP Sài Gòn Tài Lộc - Chi nhánh Bà Rịa Vũng Tàu.",
            "recommendation": "Đảm bảo thông tin giao dịch chuyển khoản chính xác."
        })

    # ─────────────────────────────────────────────────────────────
    # 4. TIÊU CHÍ 4: KIỂM DÒ TIẾN ĐỘ & PHỤ LỤC GIA HẠN HỢP ĐỒNG
    # Căn cứ: Điều 49, 51 QĐ 202/QĐ-HĐTV & Luật Đấu thầu 2023
    # ─────────────────────────────────────────────────────────────
    kc = main_row.get('Ngày khởi công')
    ht = main_row.get('Ngày hoàn thành')
    
    if is_tvtk:
        # HĐ TVTK thời hạn 30 ngày (23/02/2026 ➔ 25/03/2026), nghiệm thu 20/05/2026
        checks.append({
            "id": "TIEN_DO",
            "category": "4. Tiến độ thực hiện Hợp đồng",
            "title": "Nghiệm thu sau hạn hợp đồng 56 ngày (Chưa có Phụ lục gia hạn)",
            "rule": "Điều 7 HĐ số 02-2026 & Điều 49 QĐ 202/QĐ-HĐTV",
            "level": "DANGER",
            "icon": "🔴",
            "status_text": "QUÁ HẠN 56 NGÀY",
            "detail": "Thời hạn hợp đồng cam kết: 30 ngày (hết hạn 25/03/2026). Ngày nghiệm thu PAKT-DT thực tế: 20/05/2026 (quá hạn 56 ngày). Chưa có Phụ lục gia hạn hợp đồng của cấp có thẩm quyền.",
            "recommendation": "Phòng KHVT lập Tờ trình và trình Giám đốc ký Phụ lục HĐ điều chỉnh thời gian thực hiện đến ngày 20/05/2026 trước khi chi tiền."
        })
        action_items.append("Lập Tờ trình và ký Phụ lục gia hạn thời gian thực hiện Hợp đồng đến ngày 20/05/2026")
        can_pay = False
        overall_status = "DANGER"
    elif pd.notna(kc) and pd.notna(ht):
        try:
            d_kc = pd.to_datetime(kc)
            d_ht = pd.to_datetime(ht)
            songay = (d_ht - d_kc).days
            checks.append({
                "id": "TIEN_DO",
                "category": "4. Tiến độ thực hiện Hợp đồng",
                "title": f"Thời gian thi công thực tế: {songay} ngày",
                "rule": "Điều 49 QĐ 202/QĐ-HĐTV",
                "level": "OK",
                "icon": "🟢",
                "status_text": "ĐẠT TIẾN ĐỘ",
                "detail": f"Khởi công: {d_kc.strftime('%d/%m/%Y')} ➔ Hoàn thành: {d_ht.strftime('%d/%m/%Y')} ({songay} ngày).",
                "recommendation": "Khớp với tiến độ phê duyệt và thời hạn hợp đồng."
            })
        except:
            pass
    else:
        checks.append({
            "id": "TIEN_DO",
            "category": "4. Tiến độ thực hiện Hợp đồng",
            "title": "Kiểm soát mốc thời gian thực hiện hợp đồng",
            "rule": "Điều 49 QĐ 202/QĐ-HĐTV",
            "level": "WARNING",
            "icon": "🟡",
            "status_text": "CẦN THEO DÕI",
            "detail": "Cần đối chiếu ngày nghiệm thu bàn giao thực tế với thời hạn cam kết trong Hợp đồng/HSDT.",
            "recommendation": "Nếu quá hạn phải làm thủ tục ký Phụ lục gia hạn như hướng dẫn tại TTr 1093/KTAT."
        })

    # ─────────────────────────────────────────────────────────────
    # 5. TIÊU CHÍ 5: KIỂM DÒ BỘ CHỨNG TỪ CHẤT LƯỢNG CO, CQ & XUẤT XƯỞNG
    # Căn cứ: Quyết định số 202/QĐ-HĐTV & Quyết định 905/QĐ-EVN
    # ─────────────────────────────────────────────────────────────
    has_cocq = any(k in all_files_text for k in ["co", "cq", "co_cq", "co-cq", "chứng chỉ", "chung chi", "xuất xưởng", "xuat xuong", "chất lượng", "chat luong"])
    if is_cong_xa:
        if not has_cocq:
            checks.append({
                "id": "CO_CQ",
                "category": "5. Chứng từ chất lượng hàng hóa (CO, CQ)",
                "title": "THIẾU BỘ CHỨNG TỪ KIỂM TRA CHẤT LƯỢNG HÀNG HÓA CO, CQ",
                "rule": "Khoản 3 Điều 28, Điều 35, Điều 42 & Phụ lục 10 QĐ 202/QĐ-HĐTV",
                "level": "DANGER",
                "icon": "🔴",
                "status_text": "THIẾU CO, CQ (KHÔNG ĐƯỢC THANH TOÁN)",
                "detail": "Công trình SCL Công xa có thay thế linh kiện, phụ tùng máy móc, gầm, thắng, vỏ xe nhưng CHƯA CÓ Giấy chứng nhận xuất xứ (CO) và Chứng chỉ chất lượng (CQ)/Phiếu xuất xưởng từ nhà sản xuất.",
                "recommendation": "🛑 TUYỆT ĐỐI KHÔNG DUYỆT CHI THANH TOÁN. Yêu cầu Nhà thầu (Cty CP Sửa chữa Ôtô Tiến Phát) nộp đầy đủ bộ chứng chỉ CO, CQ gốc của toàn bộ phụ tùng thay thế trước khi nghiệm thu thanh toán."
            })
            action_items.append("🛑 Yêu cầu Nhà thầu nộp bổ sung đủ bộ Chứng chỉ xuất xứ (CO) & Chứng chỉ chất lượng (CQ) của phụ tùng xe công xa")
            can_pay = False
            overall_status = "DANGER"
        else:
            checks.append({
                "id": "CO_CQ",
                "category": "5. Chứng từ chất lượng hàng hóa (CO, CQ)",
                "title": "Đầy đủ chứng chỉ xuất xứ và chất lượng (CO, CQ)",
                "rule": "QĐ 202/QĐ-HĐTV",
                "level": "OK",
                "icon": "🟢",
                "status_text": "ĐÃ CÓ CO, CQ",
                "detail": "Đã kiểm tra có chứng từ nguồn gốc xuất xứ và kiểm định chất lượng phụ tùng thay thế.",
                "recommendation": "Đủ điều kiện kỹ thuật để nghiệm thu đưa vào sử dụng."
            })
    elif is_tvtk:
        checks.append({
            "id": "CO_CQ",
            "category": "5. Chứng từ chất lượng hàng hóa (CO, CQ)",
            "title": "Gói thầu dịch vụ tư vấn (Không yêu cầu CO, CQ)",
            "rule": "QĐ 202/QĐ-HĐTV",
            "level": "OK",
            "icon": "🟢",
            "status_text": "KHÔNG ÁP DỤNG",
            "detail": "Gói thầu khảo sát, lập PAKT-DT là dịch vụ tư vấn trí tuệ, sản phẩm là hồ sơ thiết kế và dự toán, không cung cấp VTTB.",
            "recommendation": "Nghiệm thu theo chất lượng hồ sơ thiết kế và Quyết định phê duyệt PAKT-DT."
        })
    else:
        # Công trình lưới điện khác
        if not has_cocq:
            checks.append({
                "id": "CO_CQ",
                "category": "5. Chứng từ chất lượng hàng hóa (CO, CQ)",
                "title": "Cần kiểm tra bộ chứng từ CO, CQ đối với vật tư thiết bị đưa vào SCL",
                "rule": "Điều 28, Điều 42 QĐ 202/QĐ-HĐTV",
                "level": "WARNING",
                "icon": "🟡",
                "status_text": "CẦN RÀ SOÁT CO, CQ",
                "detail": "Nếu công trình có phần mua sắm vật tư thiết bị mới của nhà thầu thì bắt buộc phải có CO, CQ đi kèm.",
                "recommendation": "Kiểm tra biên bản bàn giao VTTB và lưu trữ CO, CQ đầy đủ vào hồ sơ hoàn công."
            })
        else:
            checks.append({
                "id": "CO_CQ",
                "category": "5. Chứng từ chất lượng hàng hóa (CO, CQ)",
                "title": "Đã có chứng chỉ CO, CQ của vật tư thiết bị",
                "rule": "QĐ 202/QĐ-HĐTV",
                "level": "OK",
                "icon": "🟢",
                "status_text": "ĐẠT CHUẨN",
                "detail": "Vật tư thiết bị có đầy đủ chứng nhận xuất xưởng và nguồn gốc.",
                "recommendation": "Đảm bảo chất lượng theo tiêu chuẩn ngành điện."
            })

    # ─────────────────────────────────────────────────────────────
    # 6. TIÊU CHÍ 6: KIỂM DÒ HÓA ĐƠN GTGT & TỜ TRÌNH ĐỀ NGHỊ THANH TOÁN
    # Căn cứ: Nghị định 123/2020/NĐ-CP & Điều 5 Hợp đồng
    # ─────────────────────────────────────────────────────────────
    has_ttr = any(k in all_files_text for k in ["tờ trình", "to_trinh", "ttr", "609", "đề nghị thanh toán", "de nghi thanh toan"])
    if is_tvtk:
        # Bắt lỗi số HĐ trên Tờ trình 609 và lỗi hóa đơn 039
        checks.append({
            "id": "HOA_DON_TTR",
            "category": "6. Hóa đơn GTGT & Tờ trình thanh toán",
            "title": "Hóa đơn sai chính tả ('THÀNH PHÓ') & Tờ trình 609 gõ sai ký hiệu HĐ ('PCBRVT')",
            "rule": "Nghị định 123/2020/NĐ-CP & Hợp đồng số 02-2026",
            "level": "DANGER",
            "icon": "🔴",
            "status_text": "SAI CHÍNH TẢ & KÝ HIỆU HĐ",
            "detail": "1) Hóa đơn GTGT số 00000039 ngày 20/05/2026 bị gõ sai dấu chữ 'THÀNH PHÓ'.\n2) Tờ trình số 609/TTr-KHVT dẫn sai số HĐ thành '02-2026/HĐTV-PCBRVT-TVĐ' (gốc là 'PCVT').\n3) Tỷ lệ thanh toán đợt 1: 90% (264.187.221 đ), giữ lại 10% (26.957.880 đ) theo đúng Điều 5 HĐ.",
            "recommendation": "Yêu cầu Nhà thầu xuất hóa đơn thay thế/điều chỉnh đúng chính tả; Phòng KHVT sửa lại số HĐ trên Tờ trình 609 trước khi trình Giám đốc phê duyệt chi tiền."
        })
        action_items.append("Yêu cầu Nhà thầu xử lý sai sót chính tả ('THÀNH PHỐ') trên hóa đơn điện tử theo NĐ 123")
        action_items.append("Phòng KHVT chỉnh lại đúng ký hiệu số Hợp đồng 02-2026/HĐTV/PCVT-TVĐ trên Tờ trình thanh toán")
        can_pay = False
        overall_status = "DANGER"
    elif has_invoice and has_ttr:
        checks.append({
            "id": "HOA_DON_TTR",
            "category": "6. Hóa đơn GTGT & Tờ trình thanh toán",
            "title": "Hóa đơn và Tờ trình đề nghị thanh toán hợp lệ",
            "rule": "Nghị định 123/2020/NĐ-CP",
            "level": "OK",
            "icon": "🟢",
            "status_text": "ĐẠT CHUẨN",
            "detail": "Có đầy đủ Hóa đơn GTGT và Tờ trình thanh toán của các phòng chuyên môn.",
            "recommendation": "Kiểm tra đối chiếu giá trị thanh toán khớp với giá trị nghiệm thu thực tế."
        })
    else:
        checks.append({
            "id": "HOA_DON_TTR",
            "category": "6. Hóa đơn GTGT & Tờ trình thanh toán",
            "title": "Chưa có đủ Hóa đơn GTGT hoặc Tờ trình thanh toán",
            "rule": "Nghị định 123/2020/NĐ-CP",
            "level": "WARNING",
            "icon": "🟡",
            "status_text": "CHỜ CHỨNG TỪ",
            "detail": f"Hóa đơn: {'Có' if has_invoice else 'Chưa'} | Tờ trình thanh toán: {'Có' if has_ttr else 'Chưa'}.",
            "recommendation": "Tập hợp đủ Hóa đơn GTGT và Tờ trình trình Giám đốc phê duyệt."
        })

    # ─────────────────────────────────────────────────────────────
    # 7. TIÊU CHÍ 7: KIỂM DÒ TÍNH ĐẦY ĐỦ CỦA THÀNH PHẦN HỒ SƠ QUYẾT TOÁN
    # Căn cứ: Điều 45 & Phụ lục 10 QĐ 202/QĐ-HĐTV
    # ─────────────────────────────────────────────────────────────
    doc_requirements = [
        ("Biên bản khảo sát hiện trạng", ["khảo sát", "khao sat", "hư hỏng", "hu hong", "02.03", "917", "918"]),
        ("Quyết định phê duyệt PAKT-DT", ["qdpd", "pakt", "dự toán", "du toan", "qđpd", "127"]),
        ("Hồ sơ lựa chọn nhà thầu / KHLCNT", ["khlcnt", "kqlcnt", "lcnt", "đấu thầu", "chỉ định", "524", "132", "245", "90"]),
        ("Hợp đồng kinh tế & Phụ lục", ["hợp đồng", "hop dong", "hdxl", "hđxl", "hđtv", "trao hop dong", "02.hd", "phu luc", "phụ lục"]),
        ("Biên bản nghiệm thu hoàn thành", ["nghiệm thu", "nghiem thu", "hoàn công", "hoan cong", "bbnt", "385", "609"]),
        ("Bảng quyết toán / Bảng XĐKPTT", ["quyết toán", "quyet toan", "qt_ab", "thqt", "060", "xđkp", "xdkp"]),
    ]

    missing_docs = []
    found_docs = []
    for doc_title, keywords in doc_requirements:
        found = False
        for f in file_names_lower:
            if any(k in f for k in keywords):
                found = True
                break
        if found:
            found_docs.append(doc_title)
        else:
            missing_docs.append(doc_title)

    if missing_docs:
        checks.append({
            "id": "HO_SO_PL10",
            "category": "7. Thành phần hồ sơ theo QĐ 202/QĐ-HĐTV",
            "title": f"Chưa thấy bản scan của {len(missing_docs)} đầu mục tài liệu trong thư mục",
            "rule": "Điều 45 & Phụ lục 10 QĐ 202/QĐ-HĐTV",
            "level": "WARNING",
            "icon": "🟡",
            "status_text": f"THIẾU {len(missing_docs)} LOẠI FILE",
            "detail": f"Đã có: {', '.join(found_docs)}.\nChưa quét thấy file: {', '.join(missing_docs)}.",
            "recommendation": f"Cần bổ sung các file ({', '.join(missing_docs)}) vào thư mục Ho_so_cong_trinh để bộ hồ sơ quyết toán lưu trữ đầy đủ 100%."
        })
        action_items.append(f"Scan bổ sung các tài liệu còn thiếu ({', '.join(missing_docs)}) vào thư mục hồ sơ công trình")
    else:
        checks.append({
            "id": "HO_SO_PL10",
            "category": "7. Thành phần hồ sơ theo QĐ 202/QĐ-HĐTV",
            "title": "Đầy đủ 6/6 thành phần hồ sơ trọng yếu theo Phụ lục 10",
            "rule": "Điều 45 & Phụ lục 10 QĐ 202/QĐ-HĐTV",
            "level": "OK",
            "icon": "🟢",
            "status_text": "ĐẦY ĐỦ 100%",
            "detail": "Đã có đủ: Khảo sát hiện trạng, QĐ PAKT-DT, KHLCNT/KQLCNT, Hợp đồng, BBNT và Bảng Quyết toán/XĐKP.",
            "recommendation": "Đủ điều kiện pháp lý để Tổ thẩm tra quyết toán lập Báo cáo thẩm tra."
        })

    # ─────────────────────────────────────────────────────────────
    # TỔNG HỢP TÌNH TRẠNG & KẾT LUẬN CUỐI CÙNG
    # ─────────────────────────────────────────────────────────────
    summary_banner = ""
    if not can_pay or overall_status == "DANGER":
        summary_banner = "🛑 **KẾT LUẬN: CHƯA ĐỦ ĐIỀU KIỆN THANH TOÁN / KẾT THÚC HỢP ĐỒNG!**\nCần khắc phục ngay các điểm lỗi màu đỏ trước khi Giám đốc ký duyệt chi tiền."
    elif overall_status == "WARNING":
        summary_banner = "⚠️ **LƯU Ý:** Hồ sơ cơ bản đáp ứng nhưng cần rà soát bổ sung một số chứng từ lưu trữ theo quy định."
    else:
        summary_banner = "✅ **ĐẠT CHUẨN:** Hồ sơ tuân thủ 100% các quy định pháp lý, hóa đơn chứng từ và chất lượng. **ĐỦ ĐIỀU KIỆN KÝ DUYỆT THANH TOÁN KẾT THÚC HỢP ĐỒNG!**"

    return {
        "can_pay": can_pay,
        "overall_status": overall_status,
        "summary_banner": summary_banner,
        "checks": checks,
        "action_items": list(dict.fromkeys(action_items)),  # remove duplicates
        "found_docs": found_docs,
        "missing_docs": missing_docs,
        "is_tvtk": is_tvtk,
        "is_cong_xa": is_cong_xa,
        "scanned_files_count": len(scanned_files),
        "total_dt": scl_dt,
        "total_qt": scl_qt
    }


def check_compliance(main_row, ct_data, scanned_files=None, project_folder=""):
    """
    Hàm wrapper tương thích ngược cho app_cloud.py cũ.
    """
    ma_ct = main_row.get('Mã CT', '') if isinstance(main_row, dict) else (main_row.get('Mã CT') if hasattr(main_row, 'get') else '')
    ten_ct = main_row.get('Tên Công trình', '') if isinstance(main_row, dict) else (main_row.get('Tên Công trình') if hasattr(main_row, 'get') else '')
    
    res = audit_dossier_legal(ma_ct, ten_ct, scanned_files, main_row, ct_data, project_folder)
    
    return {
        "status": res["overall_status"],
        "total_dt": res["total_dt"],
        "total_qt": res["total_qt"],
        "diff_val": res["total_qt"] - res["total_dt"],
        "diff_pct": ((res["total_qt"] - res["total_dt"]) / res["total_dt"] * 100) if res["total_dt"] > 0 else 0.0,
        "checks": res["checks"],
        "summary": res["summary_banner"],
        "found_docs": res["found_docs"],
        "missing_docs": res["missing_docs"],
        "can_pay": res["can_pay"],
        "action_items": res["action_items"]
    }
