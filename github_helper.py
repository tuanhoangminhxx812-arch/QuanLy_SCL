"""
github_helper.py
Cung cấp các hàm tương tác với hệ thống lưu trữ tệp (Cục bộ & GitHub API)
để lưu/đọc/xoá file tài liệu đầu vào và các file dữ liệu báo cáo.
"""
import os
import base64
import requests
import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
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


def get_local_folder(ma_ct: str, doc_type: str) -> str:
    """Tạo đường dẫn thư mục cục bộ (local) trên ổ đĩa cho loại tài liệu."""
    folder = DOC_FOLDER_MAP.get(doc_type, doc_type.replace(" ", "_"))
    return os.path.join(BASE_DIR, "data", "dau_vao", ma_ct, folder)


def folder_path(ma_ct: str, doc_type: str) -> str:
    """Tạo đường dẫn thư mục trên GitHub cho loại tài liệu."""
    folder = DOC_FOLDER_MAP.get(doc_type, doc_type.replace(" ", "_"))
    return f"{BASE_PATH}/{ma_ct}/{folder}"


def _token():
    """Lấy GitHub Token từ Streamlit secrets."""
    try:
        t = st.secrets["GITHUB_TOKEN"]
        if isinstance(t, dict):
            return t.get("GITHUB_TOKEN", "")
        return str(t)
    except Exception:
        return ""


def _headers():
    return {
        "Authorization": f"token {_token()}",
        "Accept": "application/vnd.github.v3+json",
    }


def _is_valid_token():
    """Kiểm tra token có hợp lệ không (không phải placeholder)."""
    t = _token()
    if not t:
        return False
    if t.startswith("ghp_xxx") or len(t) < 20:
        return False
    return True


# Cache danh sách file từ GitHub - tồn tại 60 giây để giảm API calls
@st.cache_data(ttl=60, show_spinner=False)
def _cached_list_files(ma_ct: str, doc_type: str) -> list:
    """Gọi GitHub API và cache kết quả."""
    path = folder_path(ma_ct, doc_type)
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
    try:
        r = requests.get(url, headers=_headers(), params={"ref": GITHUB_BRANCH}, timeout=5)
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
        if r.status_code in [404, 401]:
            return []
    except Exception:
        pass
    return []


def gh_list_files(ma_ct: str, doc_type: str) -> list:
    """
    Trả về list các file trong thư mục tài liệu của công trình.
    Tự động quét thư mục cục bộ (local) và kết hợp với dữ liệu GitHub API nếu có token.
    Mỗi phần tử: {"name": str, "path": str, "sha": str, "download_url": str, "local_path": str, "is_local": bool}
    """
    files_dict = {}

    # 1. Quét thư mục cục bộ
    loc_dir = get_local_folder(ma_ct, doc_type)
    if os.path.exists(loc_dir) and os.path.isdir(loc_dir):
        for fname in os.listdir(loc_dir):
            fpath = os.path.join(loc_dir, fname)
            if os.path.isfile(fpath):
                folder = DOC_FOLDER_MAP.get(doc_type, doc_type.replace(" ", "_"))
                rel_path = f"{BASE_PATH}/{ma_ct}/{folder}/{fname}"
                files_dict[fname] = {
                    "name": fname,
                    "path": rel_path,
                    "local_path": fpath,
                    "sha": "",
                    "download_url": f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{rel_path}",
                    "size": os.path.getsize(fpath),
                    "is_local": True
                }

    # 2. Nếu có token GitHub hợp lệ, lấy thêm danh sách từ GitHub API
    if _is_valid_token():
        gh_files = _cached_list_files(ma_ct, doc_type)
        for gf in gh_files:
            fname = gf["name"]
            if fname not in files_dict:
                files_dict[fname] = {
                    "name": fname,
                    "path": gf["path"],
                    "local_path": "",
                    "sha": gf.get("sha", ""),
                    "download_url": gf.get("download_url", ""),
                    "size": 0,
                    "is_local": False
                }
            else:
                files_dict[fname]["sha"] = gf.get("sha", "")

    return list(files_dict.values())


def gh_upload_file(ma_ct: str, doc_type: str, filename: str, content_bytes: bytes) -> bool:
    """
    Lưu file vào thư mục cục bộ (local) và upload lên GitHub nếu có token.
    Trả về True nếu lưu thành công cục bộ hoặc GitHub.
    """
    local_saved = False
    try:
        loc_dir = get_local_folder(ma_ct, doc_type)
        os.makedirs(loc_dir, exist_ok=True)
        target_path = os.path.join(loc_dir, filename)
        with open(target_path, "wb") as f:
            f.write(content_bytes)
        local_saved = True
    except Exception as e:
        print(f"Lỗi khi lưu file cục bộ: {e}")

    gh_saved = False
    if _is_valid_token():
        path = f"{folder_path(ma_ct, doc_type)}/{filename}"
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
        sha = None
        try:
            r = requests.get(url, headers=_headers(), params={"ref": GITHUB_BRANCH}, timeout=5)
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
            if r.status_code in [200, 201]:
                gh_saved = True
        except Exception:
            pass

    try:
        _cached_list_files.clear()
    except Exception:
        pass

    return local_saved or gh_saved


def gh_delete_file(path: str, sha: str = "", ma_ct: str = "", doc_type: str = "", filename: str = "") -> bool:
    """
    Xoá file khỏi thư mục cục bộ và xoá trên GitHub nếu có token.
    """
    local_deleted = False

    # 1. Xóa theo path tương đối
    if path:
        local_target = os.path.join(BASE_DIR, path.replace("/", os.sep))
        if os.path.exists(local_target) and os.path.isfile(local_target):
            try:
                os.remove(local_target)
                local_deleted = True
            except Exception:
                pass

    # 2. Xóa theo ma_ct, doc_type, filename nếu path chưa xóa được
    if not local_deleted and ma_ct and doc_type and filename:
        alt_path = os.path.join(get_local_folder(ma_ct, doc_type), filename)
        if os.path.exists(alt_path) and os.path.isfile(alt_path):
            try:
                os.remove(alt_path)
                local_deleted = True
            except Exception:
                pass

    # 3. Xoá trên GitHub nếu có token
    gh_deleted = False
    if _is_valid_token() and path:
        if not sha:
            try:
                url_get = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
                r_get = requests.get(url_get, headers=_headers(), params={"ref": GITHUB_BRANCH}, timeout=5)
                if r_get.status_code == 200:
                    sha = r_get.json().get("sha", "")
            except Exception:
                pass

        if sha:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"
            payload = {
                "message": f"Delete: {path}",
                "sha": sha,
                "branch": GITHUB_BRANCH,
            }
            try:
                r = requests.delete(url, headers=_headers(), json=payload, timeout=10)
                if r.status_code == 200:
                    gh_deleted = True
            except Exception:
                pass

    try:
        _cached_list_files.clear()
    except Exception:
        pass

    return local_deleted or gh_deleted


def gh_upload_root_file(filename: str, content_bytes: bytes) -> bool:
    """Lưu file vào thư mục gốc cục bộ và cập nhật lên GitHub nếu có token."""
    local_saved = False
    try:
        target_path = os.path.join(BASE_DIR, filename)
        with open(target_path, "wb") as f:
            f.write(content_bytes)
        local_saved = True
    except Exception as e:
        print(f"Lỗi khi lưu file gốc cục bộ: {e}")

    gh_saved = False
    if _is_valid_token():
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filename}"
        sha = None
        try:
            r = requests.get(url, headers=_headers(), params={"ref": GITHUB_BRANCH}, timeout=5)
            if r.status_code == 200:
                sha = r.json().get("sha")
        except Exception:
            pass

        payload = {
            "message": f"Cập nhật {filename}",
            "content": base64.b64encode(content_bytes).decode(),
            "branch": GITHUB_BRANCH,
        }
        if sha:
            payload["sha"] = sha

        try:
            r = requests.put(url, headers=_headers(), json=payload, timeout=30)
            if r.status_code in [200, 201]:
                gh_saved = True
        except Exception:
            pass

    return local_saved or gh_saved


def has_token() -> bool:
    """Kiểm tra xem GitHub Token đã được cài đặt và hợp lệ chưa."""
    return _is_valid_token()
