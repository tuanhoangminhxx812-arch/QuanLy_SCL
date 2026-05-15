"""
github_helper.py
Cung cấp các hàm tương tác với GitHub API để lưu/đọc/xoá file tài liệu đầu vào.
"""
import base64
import requests
import streamlit as st

GITHUB_REPO = "tuanhoangminhxx812-arch/QuanLy_SCL"
GITHUB_BRANCH = "main"
BASE_PATH = "data/dau_vao"

# Mapping tên loại tài liệu → tên thư mục an toàn (ASCII)
DOC_FOLDER_MAP = {
    "QĐ kế hoạch vốn SCL Công ty":       "01_QD_ke_hoach_von",
    "PAKT-DT được duyệt":                 "02_PAKT_DT",
    "Kế hoạch Đấu thầu":                  "03_KH_dau_thau",
    "Kết quả Đấu thầu":                   "04_KQ_dau_thau",
    "Hợp Đồng":                           "05_Hop_dong",
    "Các Biên Bản Nghiệm Thu":            "06_Bien_ban_nghiem_thu",
    "Quyết toán A-B":                     "07_Quyet_toan_AB",
    "Bảng Tổng Hợp (CT SCL của TCT)":    "08_Bang_tong_hop",
    "PM_092 (ERP)":                       "09_PM092",
}


def _token():
    """Lấy GitHub Token từ Streamlit secrets."""
    try:
        return st.secrets["GITHUB_TOKEN"]
    except Exception:
        return ""


def _headers():
    return {
        "Authorization": f"token {_token()}",
        "Accept": "application/vnd.github.v3+json",
    }


def folder_path(ma_ct: str, doc_type: str) -> str:
    """Tạo đường dẫn thư mục trên GitHub cho loại tài liệu."""
    folder = DOC_FOLDER_MAP.get(doc_type, doc_type.replace(" ", "_"))
    return f"{BASE_PATH}/{ma_ct}/{folder}"


def gh_list_files(ma_ct: str, doc_type: str) -> list:
    """
    Trả về list các file trong thư mục tài liệu của công trình.
    Mỗi phần tử: {"name": str, "path": str, "sha": str, "download_url": str}
    """
    path = folder_path(ma_ct, doc_type)
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
    try:
        r = requests.get(url, headers=_headers(), params={"ref": GITHUB_BRANCH}, timeout=10)
        if r.status_code == 200:
            return [
                {
                    "name": item["name"],
                    "path": item["path"],
                    "sha": item["sha"],
                    "download_url": item["download_url"],
                }
                for item in r.json()
                if item["type"] == "file"
            ]
    except Exception:
        pass
    return []


def gh_upload_file(ma_ct: str, doc_type: str, filename: str, content_bytes: bytes) -> bool:
    """
    Upload hoặc cập nhật một file lên GitHub.
    Trả về True nếu thành công.
    """
    path = f"{folder_path(ma_ct, doc_type)}/{filename}"
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"

    # Kiểm tra file đã tồn tại chưa (cần SHA để update)
    sha = None
    try:
        r = requests.get(url, headers=_headers(), params={"ref": GITHUB_BRANCH}, timeout=10)
        if r.status_code == 200:
            sha = r.json().get("sha")
    except Exception:
        pass

    payload = {
        "message": f"Upload: {ma_ct}/{doc_type}/{filename}",
        "content": base64.b64encode(content_bytes).decode(),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha

    try:
        r = requests.put(url, headers=_headers(), json=payload, timeout=30)
        return r.status_code in [200, 201]
    except Exception:
        return False


def gh_delete_file(path: str, sha: str) -> bool:
    """
    Xoá một file khỏi GitHub theo đường dẫn và SHA.
    Trả về True nếu thành công.
    """
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
    payload = {
        "message": f"Delete: {path}",
        "sha": sha,
        "branch": GITHUB_BRANCH,
    }
    try:
        r = requests.delete(url, headers=_headers(), json=payload, timeout=10)
        return r.status_code == 200
    except Exception:
        return False


def has_token() -> bool:
    """Kiểm tra xem GitHub Token đã được cài đặt chưa."""
    return bool(_token())
