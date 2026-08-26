"""
legal_checker.py
Module kiểm tra tuân thủ quy định và phân tích pháp lý hồ sơ quyết toán SCL
Căn cứ theo:
1. Quyết định số 202/QĐ-HĐTV ngày 31/12/2025 của Hội đồng thành viên Tổng công ty Điện lực TP.HCM
   về Quy định thực hiện công tác sửa chữa lớn tài sản trong EVNHCMC.
2. Quyết định số 905/QĐ-EVN ngày 17/06/2025 của EVN về Quản lý kỹ thuật.
3. Thông tư số 45/2013/TT-BTC & Thông tư 99/TT-BTC về quản lý tài sản và kế toán.
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


def check_compliance(main_row, ct_data, scanned_files=None, project_folder=""):
    """
    Thực hiện kiểm tra tính hợp lệ và tuân thủ quy định của bộ hồ sơ quyết toán SCL.
    Trả về dict chứa danh sách các điểm kiểm tra và nhận định pháp lý.
    """
    if scanned_files is None:
        scanned_files = []

    checks = []
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

    diff_val = scl_qt - scl_dt
    diff_pct = (diff_val / scl_dt * 100) if scl_dt > 0 else 0.0

    # ─────────────────────────────────────────────────────────────
    # TIÊU CHÍ 1: Tổng giá trị quyết toán so với Dự toán được duyệt
    # Căn cứ: Điều 43 QĐ 202/QĐ-HĐTV
    # ─────────────────────────────────────────────────────────────
    if scl_dt > 0 and scl_qt > 0:
        if scl_qt > scl_dt:
            overall_status = "DANGER"
            checks.append({
                "category": "Giá trị quyết toán",
                "title": "Tổng giá trị quyết toán vượt dự toán được duyệt",
                "rule": "Điều 43 QĐ 202/QĐ-HĐTV: Quyết toán phải trên cơ sở dự toán được duyệt",
                "level": "DANGER",
                "icon": "🔴",
                "detail": f"Giá trị quyết toán ({_fmt_money(scl_qt)} đ) vượt dự toán duyệt ({_fmt_money(scl_dt)} đ) là {_fmt_money(diff_val)} đ (+{diff_pct:.2f}%).",
                "recommendation": "Theo Điều 43, nghiêm cấm quyết toán vượt dự toán mà không có Quyết định phê duyệt điều chỉnh dự toán/kế hoạch của cấp có thẩm quyền trước khi quyết toán."
            })
        elif scl_qt < scl_dt:
            tiet_kiem = scl_dt - scl_qt
            checks.append({
                "category": "Giá trị quyết toán",
                "title": "Quyết toán trong hạn mức dự toán (Tiết kiệm chi phí)",
                "rule": "Điều 43 QĐ 202/QĐ-HĐTV",
                "level": "OK",
                "icon": "🟢",
                "detail": f"Quyết toán ({_fmt_money(scl_qt)} đ) tiết kiệm so với dự toán ({_fmt_money(scl_dt)} đ) là {_fmt_money(tiet_kiem)} đ (-{abs(diff_pct):.2f}%).",
                "recommendation": "Hồ sơ tuân thủ quy định về giá trị, không làm phát sinh tăng vốn SCL."
            })
        else:
            checks.append({
                "category": "Giá trị quyết toán",
                "title": "Giá trị quyết toán bằng đúng dự toán được duyệt",
                "rule": "Điều 43 QĐ 202/QĐ-HĐTV",
                "level": "OK",
                "icon": "🟢",
                "detail": f"Giá trị quyết toán bằng đúng 100% dự toán ({_fmt_money(scl_qt)} đ).",
                "recommendation": "Đảm bảo tính chính xác và khớp số liệu dự toán."
            })
    else:
        checks.append({
            "category": "Giá trị quyết toán",
            "title": "Chưa có đủ số liệu Dự toán hoặc Quyết toán",
            "rule": "Điều 43 QĐ 202/QĐ-HĐTV",
            "level": "WARNING",
            "icon": "🟡",
            "detail": f"Dự toán: {_fmt_money(scl_dt)} đ | Quyết toán: {_fmt_money(scl_qt)} đ.",
            "recommendation": "Cần quét hoặc cập nhật đầy đủ số liệu Dự toán và Quyết toán A-B vào hệ thống."
        })

    # ─────────────────────────────────────────────────────────────
    # TIÊU CHÍ 2: Kiểm tra Giá trị vật tư, thiết bị thu hồi (VTTBTH)
    # Căn cứ: Khoản 4 Điều 28 & Điều 42 QĐ 202/QĐ-HĐTV
    # ─────────────────────────────────────────────────────────────
    vt_th_dt = _safe_num(bd.get('F', {}).get('dt', 0))
    vt_th_qt = _safe_num(bd.get('F', {}).get('qt', 0))
    if vt_th_dt == 0:
        vt_th_dt = _safe_num(main_row.get('Giá trị VTTH', 0))

    if vt_th_dt > 0:
        if vt_th_qt > 0:
            checks.append({
                "category": "Vật tư thu hồi",
                "title": "Đã thực hiện giảm trừ giá trị vật tư thu hồi",
                "rule": "Khoản 4 Điều 28 & Điều 42 QĐ 202/QĐ-HĐTV",
                "level": "OK",
                "icon": "🟢",
                "detail": f"Dự toán thu hồi: {_fmt_money(vt_th_dt)} đ, Quyết toán giảm trừ: {_fmt_money(vt_th_qt)} đ.",
                "recommendation": "Đã tuân thủ quy định hạch toán giảm giá trị công trình SCL theo quy định."
            })
        else:
            if overall_status != "DANGER":
                overall_status = "WARNING"
            checks.append({
                "category": "Vật tư thu hồi",
                "title": "Dự toán có VTTB thu hồi nhưng quyết toán chưa thể hiện giảm trừ",
                "rule": "Khoản 4 Điều 28 & Điều 42 QĐ 202/QĐ-HĐTV",
                "level": "WARNING",
                "icon": "🟡",
                "detail": f"Trong dự toán được duyệt có tính giá trị VTTB thu hồi là {_fmt_money(vt_th_dt)} đ nhưng trong quyết toán chưa thấy giảm trừ.",
                "recommendation": "Theo Khoản 4 Điều 28, 'Giá trị vật tư thu hồi được hạch toán giảm giá trị công trình SCL'. Cần kiểm tra biên bản thu hồi, phiếu nhập kho vật tư thu hồi và lập bảng giảm trừ quyết toán."
            })
    else:
        checks.append({
            "category": "Vật tư thu hồi",
            "title": "Công trình không có vật tư thu hồi hoặc giá trị thu hồi bằng 0",
            "rule": "Điều 42 QĐ 202/QĐ-HĐTV",
            "level": "OK",
            "icon": "🟢",
            "detail": "Không phát sinh vật tư thiết bị thu hồi trong dự toán.",
            "recommendation": "Phù hợp với tính chất công trình."
        })

    # ─────────────────────────────────────────────────────────────
    # TIÊU CHÍ 3: Kiểm tra Chi phí dự phòng trong quyết toán
    # Căn cứ: QĐ 202/QĐ-HĐTV & Thông tư Bộ Xây dựng
    # ─────────────────────────────────────────────────────────────
    dp_dt = _safe_num(bd.get('D', {}).get('dt', 0))
    dp_qt = _safe_num(bd.get('D', {}).get('qt', 0))
    if dp_qt > 0:
        checks.append({
            "category": "Chi phí dự phòng",
            "title": "Có quyết toán chi phí dự phòng",
            "rule": "Quy định quản lý chi phí SCL",
            "level": "WARNING",
            "icon": "🟡",
            "detail": f"Quyết toán chi phí dự phòng: {_fmt_money(dp_qt)} đ.",
            "recommendation": "Chi phí dự phòng chỉ được quyết toán khi có khối lượng phát sinh thực tế được phê duyệt bằng văn bản của cấp có thẩm quyền."
        })
    else:
        checks.append({
            "category": "Chi phí dự phòng",
            "title": "Chi phí dự phòng quyết toán bằng 0 (Đúng quy định)",
            "rule": "Quy định quản lý chi phí SCL",
            "level": "OK",
            "icon": "🟢",
            "detail": f"Dự toán dự phòng: {_fmt_money(dp_dt)} đ ➔ Quyết toán: 0 đ.",
            "recommendation": "Quyết toán theo chi phí thực tế phát sinh hợp pháp, không quyết toán dự phòng thừa."
        })

    # ─────────────────────────────────────────────────────────────
    # TIÊU CHÍ 4: Kiểm tra Tiến độ thi công so với Hợp đồng / Kế hoạch
    # Căn cứ: Điều 49, 51 QĐ 202/QĐ-HĐTV
    # ─────────────────────────────────────────────────────────────
    kc = main_row.get('Ngày khởi công')
    ht = main_row.get('Ngày hoàn thành')
    
    if pd.notna(kc) and pd.notna(ht):
        try:
            d_kc = pd.to_datetime(kc)
            d_ht = pd.to_datetime(ht)
            songay = (d_ht - d_kc).days
            checks.append({
                "category": "Tiến độ thực hiện",
                "title": f"Thời gian thi công thực tế: {songay} ngày",
                "rule": "Điều 49 QĐ 202/QĐ-HĐTV",
                "level": "OK",
                "icon": "🟢",
                "detail": f"Khởi công: {d_kc.strftime('%d/%m/%Y')} ➔ Hoàn thành: {d_ht.strftime('%d/%m/%Y')} ({songay} ngày).",
                "recommendation": "Cần kiểm tra đối chiếu với tiến độ cam kết trong Hợp đồng để đảm bảo không vi phạm thời hạn."
            })
        except:
            pass
    elif pd.notna(kc) and pd.isna(ht):
        checks.append({
            "category": "Tiến độ thực hiện",
            "title": "Chưa có ngày hoàn thành công trình",
            "rule": "Điều 49 QĐ 202/QĐ-HĐTV",
            "level": "WARNING",
            "icon": "🟡",
            "detail": f"Đã có ngày khởi công ({pd.to_datetime(kc).strftime('%d/%m/%Y') if pd.notna(kc) else ''}) nhưng chưa có ngày nghiệm thu hoàn thành.",
            "recommendation": "Cần bổ sung Biên bản nghiệm thu hoàn thành để xác định thời điểm kết thúc thi công và chốt số liệu quyết toán."
        })

    # ─────────────────────────────────────────────────────────────
    # TIÊU CHÍ 5: Kiểm tra tính đầy đủ của thành phần hồ sơ theo Điều 45 & Phụ lục 10
    # ─────────────────────────────────────────────────────────────
    file_names_lower = [os.path.basename(f).lower() for f in scanned_files]
    all_text_combined = " ".join(file_names_lower)

    doc_requirements = [
        ("Biên bản khảo sát hiện trạng", ["khảo sát", "khao sat", "hư hỏng", "hu hong", "02.03"]),
        ("Quyết định phê duyệt PAKT-DT", ["qdpd", "pakt", "dự toán", "du toan", "qđpd"]),
        ("Hồ sơ lựa chọn nhà thầu / KHLCNT", ["khlcnt", "kqlcnt", "lcnt", "đấu thầu", "chỉ định"]),
        ("Hợp đồng kinh tế / Thư trao thầu", ["hợp đồng", "hop dong", "hdxl", "hđxl", "trao hop dong", "trao hợp đồng"]),
        ("Biên bản nghiệm thu hoàn công", ["nghiệm thu", "nghiem thu", "hoàn công", "hoan cong", "bbnt"]),
        ("Hồ sơ / Bảng quyết toán A-B", ["quyết toán", "quyet toan", "qt_ab", "thqt", "quyết toán a-b", "bảng thqt"]),
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
        if overall_status != "DANGER":
            overall_status = "WARNING"
        checks.append({
            "category": "Thành phần hồ sơ",
            "title": f"Phát hiện thiếu {len(missing_docs)} đầu mục tài liệu theo Điều 45 & Phụ lục 10",
            "rule": "Điều 45 & Phụ lục 10 QĐ 202/QĐ-HĐTV",
            "level": "WARNING",
            "icon": "🟡",
            "detail": f"Đã có: {', '.join(found_docs)}. Chưa thấy: {', '.join(missing_docs)}.",
            "recommendation": f"Cần bổ sung các tài liệu còn thiếu ({', '.join(missing_docs)}) vào thư mục hồ sơ công trình để hoàn thiện bộ hồ sơ quyết toán hợp lệ."
        })
    else:
        checks.append({
            "category": "Thành phần hồ sơ",
            "title": "Đầy đủ các thành phần hồ sơ trọng yếu",
            "rule": "Điều 45 & Phụ lục 10 QĐ 202/QĐ-HĐTV",
            "level": "OK",
            "icon": "🟢",
            "detail": f"Thư mục đã có đủ các loại tài liệu: Khảo sát, PAKT-DT, KHLCNT, Hợp đồng, Nghiệm thu, Quyết toán.",
            "recommendation": "Đủ điều kiện pháp lý để Tổ thẩm tra tiến hành thẩm tra và trình Giám đốc phê duyệt."
        })

    # ─────────────────────────────────────────────────────────────
    # Tổng hợp nhận định
    # ─────────────────────────────────────────────────────────────
    summary_msg = []
    if overall_status == "DANGER":
        summary_msg.append("⚠️ **CẢNH BÁO PHÁP LÝ:** Hồ sơ có nội dung chưa tuân thủ quy định của QĐ 202/QĐ-HĐTV (vượt dự toán duyệt). Cần xử lý điều chỉnh dự toán trước khi phê duyệt.")
    elif overall_status == "WARNING":
        summary_msg.append("⚡ **LƯU Ý:** Hồ sơ cơ bản đáp ứng nhưng cần rà soát bổ sung một số tài liệu hoặc giảm trừ VTTB thu hồi theo quy định.")
    else:
        summary_msg.append("✅ **ĐẠT CHUẨN:** Hồ sơ tuân thủ đầy đủ quy định tại QĐ 202/QĐ-HĐTV của Tổng công ty Điện lực TP.HCM. Đủ điều kiện phê duyệt quyết toán.")

    return {
        "status": overall_status,
        "total_dt": scl_dt,
        "total_qt": scl_qt,
        "diff_val": diff_val,
        "diff_pct": diff_pct,
        "checks": checks,
        "summary": " ".join(summary_msg),
        "found_docs": found_docs,
        "missing_docs": missing_docs,
    }
