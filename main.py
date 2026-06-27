import tkinter as tk
from PIL import Image, ImageTk
import datetime
import os
from PIL import Image, ImageTk
import pyperclip
from tkinter import ttk
import shutil
import re
import requests
import msal
import base64
import time
import threading
from tkinter import ttk, messagebox
import json
import schedule
from tkcalendar import DateEntry

# ==== Khởi tạo Tkinter root trước ====
root = tk.Tk()
root.withdraw()   # Ẩn cửa sổ chính ban đầu

# ==== Thiết lập và Cấu hình Azure AD, OneDrive, đường dẫn lưu trữ và hơn thế nữa =============================================================================================================
BASE_URL = (
    "https://aeondelight-my.sharepoint.com/"
    "personal/phuc_nguyen_aeondelight_biz/"
    "Documents/PHUC/PHUC/AZURE/"
    "RMC%20DATA%20STORAGE%20V2"
)

# ============= LINK ONEDRIVE OF REPORT FORM ===============
# ===== LINK LƯU TRỮ CÁC TÀI LIỆU PDF =====
documentary_archive_url = f"{BASE_URL}/DOCUMENTARY"

# == Thông tin ID của ứng dụng Azure AD ==
CLIENT_ID = "ac4edccf-a8ee-41aa-bcc4-6603c4bebae1"
TENANT_ID = "5983a1d2-f46b-492d-a9b3-7e2f3609d20b"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_SCOPES = ["Files.Read"]
CACHE_DIR = r"D:\RMC_Assistant_v3\Cache"
CACHE_FILE = os.path.join(CACHE_DIR, "token_cache.bin")

# ============ Đường dân local trên máy tính để lưu trữ cache ================
APP_ROOT = os.path.join(os.environ["LOCALAPPDATA"],"RMC_Assistant_v3")
# ==========================================================
# CACHE
# ==========================================================
CACHE_DIR = os.path.join(APP_ROOT, "Cache")
CACHE_FILE = os.path.join(CACHE_DIR, "token_cache.bin")
# ==========================================================
# REPORT FORM
# ==========================================================
REPORT_FORM_DIR = os.path.join(APP_ROOT, "Report_Form_Cache")
AEONGMS_DIR = os.path.join(REPORT_FORM_DIR, "AEONGMS")
AEONMAXVALU_DIR = os.path.join(REPORT_FORM_DIR, "MAXVALU")
# ==========================================================
# NOTE
# ==========================================================
NOTE_ARCHIVE_DIR = os.path.join(APP_ROOT, "NOTE")
# ==========================================================
# IMAGE
# ==========================================================
IMAGE_ROOT_DIR = os.path.join(APP_ROOT, "IMAGE")
AEONGMSIMG_DIR = os.path.join(IMAGE_ROOT_DIR, "AEONGMS")
MAXVALUIMG_DIR = os.path.join(IMAGE_ROOT_DIR, "MAXVALU")

# AEON GMS IMG ARCHIVE
AEONGMS_IMAGE_LAYOUT_ARCHIVE_DIR = os.path.join(AEONGMSIMG_DIR, "LAYOUT")
AEONGMS_IMAGE_GATEWAY_ARCHIVE_DIR = os.path.join(AEONGMSIMG_DIR, "GATEWAY")
AEONGMS_IMAGE_SENSOR_ARCHIVE_DIR = os.path.join(AEONGMSIMG_DIR, "SENSOR")
AEONGMS_IMAGE_AL_ARCHIVE_DIR = os.path.join(AEONGMSIMG_DIR, "ALARM_POINTS") # Đã sửa thành ALARM_POINTS

# AEON MAXVALU IMG ARCHIVE
MAXVALU_IMAGE_LAYOUT_ARCHIVE_DIR = os.path.join(MAXVALUIMG_DIR, "LAYOUT")
MAXVALU_IMAGE_GATEWAY_ARCHIVE_DIR = os.path.join(MAXVALUIMG_DIR, "GATEWAY")
MAXVALU_IMAGE_SENSOR_ARCHIVE_DIR = os.path.join(MAXVALUIMG_DIR, "SENSOR")
MAXVALU_IMAGE_AL_ARCHIVE_DIR = os.path.join(MAXVALUIMG_DIR, "ALARM_POINTS")
# ==========================================================
# DOCUMENTARY
# ==========================================================
DOCUMENTARY_ARCHIVE_DIR = os.path.join(APP_ROOT,"DOCUMENTARY")
# ==========================================================
# METADATA
# ==========================================================
METADATA_DIR = os.path.join(APP_ROOT,"METADATA")
# ==========================================================
# TẠO THƯ MỤC NẾU CHƯA TỒN TẠI
# ==========================================================
folders = [
    APP_ROOT,
    CACHE_DIR,
    REPORT_FORM_DIR,
    AEONGMS_DIR,
    AEONMAXVALU_DIR,
    NOTE_ARCHIVE_DIR,
    IMAGE_ROOT_DIR,
    
    # Nhánh thư mục AEON GMS
    AEONGMSIMG_DIR,
    AEONGMS_IMAGE_LAYOUT_ARCHIVE_DIR,
    AEONGMS_IMAGE_GATEWAY_ARCHIVE_DIR,
    AEONGMS_IMAGE_SENSOR_ARCHIVE_DIR,
    AEONGMS_IMAGE_AL_ARCHIVE_DIR,

    # Nhánh thư mục MAXVALU
    MAXVALUIMG_DIR,
    MAXVALU_IMAGE_LAYOUT_ARCHIVE_DIR,
    MAXVALU_IMAGE_GATEWAY_ARCHIVE_DIR,
    MAXVALU_IMAGE_SENSOR_ARCHIVE_DIR,
    MAXVALU_IMAGE_AL_ARCHIVE_DIR,

    DOCUMENTARY_ARCHIVE_DIR,
    METADATA_DIR
]
for folder in folders:
    os.makedirs(folder, exist_ok=True)

print("APP_ROOT =", APP_ROOT)
# ==========================================================
# BIẾN HỆ THỐNG
# ==========================================================
CURRENT_SOURCE = "AEONGMS"
ALL_REPORT_MAPPINGS = {
    "AEONGMS": {},
    "MAXVALU": {}
}

# ==== Đăng nhập, tải file và xử lý OneDrive bằng Azure ===========================================================================
# == Cửa sổ đăng nhập Azure AD trên thiết bị mới ==
def show_device_login(flow):
    win = tk.Toplevel(root)   # ✅ gắn với root chính
    win.title("🔑 Đăng nhập Azure AD")
    win.geometry("500x300")
    win.grab_set()  # ✅ chặn tương tác ngoài login

    # Frame 1: Thông báo
    frame1 = tk.Frame(win, pady=10)
    frame1.pack(fill="x")
    tk.Label(frame1, text="Bạn đang đăng nhập trên một thiết bị mới",
             font=("Arial", 12, "bold"), fg="red").pack()

    # Frame 2: Văn bản truy cập
    frame2 = tk.Frame(win, pady=5)
    frame2.pack(fill="x")
    tk.Label(frame2, text="Truy cập vào liên kết dưới đây:", font=("Arial", 11)).pack()

    # Frame 3: Link
    frame3 = tk.Frame(win, pady=5)
    frame3.pack(fill="x")
    entry_link = tk.Entry(frame3, font=("Arial", 11), width=50)
    entry_link.insert(0, flow["verification_uri"])
    entry_link.pack(padx=10)

    # Frame 4: Văn bản nhập mã
    frame4 = tk.Frame(win, pady=5)
    frame4.pack(fill="x")
    tk.Label(frame4, text="Nhập mã sau vào trang web:", font=("Arial", 11)).pack()

    # Frame 5: Mã đăng nhập
    frame5 = tk.Frame(win, pady=5)
    frame5.pack(fill="x")
    entry_code = tk.Entry(frame5, font=("Arial", 14, "bold"), width=20, justify="center")
    entry_code.insert(0, flow["user_code"])
    entry_code.pack(padx=10)

    # Copy link và code
    def copy_link():
        win.clipboard_clear()
        win.clipboard_append(flow["verification_uri"])

    def copy_code():
        win.clipboard_clear()
        win.clipboard_append(flow["user_code"])

    btn_frame = tk.Frame(win, pady=10)
    btn_frame.pack()
    ttk.Button(btn_frame, text="📋 Copy link", command=copy_link).grid(row=0, column=0, padx=10)
    ttk.Button(btn_frame, text="📋 Copy mã", command=copy_code).grid(row=0, column=1, padx=10)

    return win

# ==== Hàm đăng nhập Azure AD ====
def authenticate():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r").read())

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()

    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if "user_code" not in flow:
            raise Exception("Không khởi tạo được Device Flow")

        win = show_device_login(flow)
        result_container = {"result": None}

        def do_login():
            result = app.acquire_token_by_device_flow(flow)
            result_container["result"] = result

        threading.Thread(target=do_login, daemon=True).start()

        def check_result():
            if result_container["result"] is not None:
                win.destroy()
            else:
                root.after(500, check_result)

        root.after(500, check_result)
        win.wait_window()
        result = result_container["result"]

    # ✅ Lưu lại cache sau khi login hoặc refresh thành công
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

    return result

# ==== Đăng nhập Azure ====
try:
    result = authenticate()

    if "access_token" not in result:
        raise Exception("Đăng nhập thất bại")

    access_token = result["access_token"]
except Exception as e:

    messagebox.showerror("Lỗi", str(e))

    root.destroy()

    exit()

# ==== FILE ĐÁNH DẤU CHẠY LẦN ĐẦU =========================================
FIRST_RUN_FILE = os.path.join(
    CACHE_DIR,
    "first_run.json"
)
def is_first_run():
    return not os.path.exists(FIRST_RUN_FILE)
def create_first_run_flag():
    now = datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    token_created = "Unknown"
    try:
        if os.path.exists(CACHE_FILE):
            ts = os.path.getmtime(CACHE_FILE)
            token_created = datetime.datetime.fromtimestamp(
                ts
            ).strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        print("❌ Token time error:", e)

    data = {
        "first_run_time": now,
        "token_cache_created": token_created
    }

    with open(
        FIRST_RUN_FILE,
        "w",
        encoding="utf-8"
    ) as f:
        json.dump(
            data,
            f,
            indent=4,
            ensure_ascii=False
        )

    print("✅ Đã tạo first_run.flag")
def first_time_auto_sync():
    try:
        print("🚀 FIRST RUN DETECTED")
        # =====================================================
        # LOGIN AZURE
        # =====================================================
        graph_session.ensure_token()
        # =====================================================
        # BUILD DATA LINK
        # =====================================================
        build_data_link()
        # =====================================================
        # AUTO SYNC
        # =====================================================
        auto_sync_all_onedrive()
        # =====================================================
        # CREATE FLAG
        # =====================================================
        create_first_run_flag()
        print("✅ FIRST RUN COMPLETED")
    except Exception as e:
        print("❌ FIRST RUN ERROR:", e)

# === Giao diện tùy chọn sync hay vào thẳng app ===
def launch_main_program():
    global DATA_LINK
    DATA_LINK = load_data_link_json()
    refresh_report_mapping()
    root.deiconify()
    root.title("RMC Assistant")
    root.state("zoomed")
    create_area_buttons()
    start_clock()

def run_sync_and_launch(startup, status_label):
    try:
        status_label.config(
            text="🔄 Đang build data link..."
        )
        root.update_idletasks()
        # =========================
        # BUILD LINK
        # =========================
        build_data_link()
        status_label.config(
            text="🔄 Đang đồng bộ dữ liệu..."
        )
        root.update_idletasks()
        # =========================
        # SYNC
        # =========================
        auto_sync_all_onedrive()
        status_label.config(
            text="✅ Đồng bộ hoàn tất"
        )
        root.update_idletasks()
        time.sleep(1)

    except Exception as e:
        messagebox.showerror(
            "Sync Error",
            str(e)
        )

    finally:
        startup.destroy()
        global DATA_LINK
        DATA_LINK = load_data_link_json()
        launch_main_program()

def skip_sync(startup):
    startup.destroy()
    global DATA_LINK
    DATA_LINK = load_data_link_json()
    refresh_report_mapping()
    launch_main_program()

# >> Sau một thời gian chương trình treo (idle) thì access_token hết hạn (thường là 1 giờ), nên không lấy được dữ liệu thường xuyên <<
# ==== Quản lý phiên làm việc với Azure Graph API ====
class GraphSession:
    def __init__(self, client_id, authority, scopes, cache_file):
        self.client_id = client_id
        self.authority = authority
        self.scopes = scopes
        self.cache_file = cache_file
        self.cache = msal.SerializableTokenCache()
        if os.path.exists(cache_file):
            self.cache.deserialize(open(cache_file, "r").read())
        self.app = msal.PublicClientApplication(client_id, authority=authority, token_cache=self.cache)
        self.account = None
        self.token = None

    def save_cache(self):
        if self.cache.has_state_changed:
            with open(self.cache_file, "w") as f:
                f.write(self.cache.serialize())

    def ensure_token(self):
        """Đảm bảo luôn có access_token hợp lệ (refresh khi cần)."""
        # Nếu đã có token thì check hạn
        if self.token and "access_token" in self.token:
            expires_at = self.token.get("expires_on")
            if expires_at:
                import time
                if int(expires_at) > int(time.time()) + 60:
                    # Token còn sống > 60s thì dùng tiếp
                    return self.token["access_token"]

        # Nếu chưa có hoặc hết hạn → thử silent refresh
        accounts = self.app.get_accounts()
        if accounts:
            self.account = accounts[0]
            self.token = self.app.acquire_token_silent(self.scopes, account=self.account)

        # Nếu vẫn chưa có token hợp lệ → login lại
        if not self.token or "access_token" not in self.token:
            flow = self.app.initiate_device_flow(scopes=self.scopes)
            if "user_code" not in flow:
                raise Exception("Không khởi tạo được Device Flow")
            win = show_device_login(flow)
            result_container = {"result": None}

            def do_login():
                result = self.app.acquire_token_by_device_flow(flow)
                result_container["result"] = result

            threading.Thread(target=do_login, daemon=True).start()

            def check_result():
                if result_container["result"] is not None:
                    win.destroy()
                else:
                    root.after(500, check_result)

            root.after(500, check_result)
            win.wait_window()
            self.token = result_container["result"]

        if not self.token or "access_token" not in self.token:
            raise Exception("Đăng nhập Azure thất bại")

        self.save_cache()
        return self.token["access_token"]

# ==== Khởi tạo session Azure ====
graph_session = GraphSession(CLIENT_ID, AUTHORITY, GRAPH_SCOPES, CACHE_FILE)

# ==== AUTO BUILD DATA LINK =========================================================
DATA_LINK_FILE = os.path.join(CACHE_DIR,"data_link.json")

# ==== Encode URL =========================================================
def encode_share_url(url):
    encoded = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    encoded = (
        encoded.rstrip("=")
        .replace("/", "_")
        .replace("+", "-")
    )
    return encoded

# ==== GET SUB FOLDER =========================================================
def get_folders_from_share_url(token, share_url):
    encoded_url = encode_share_url(share_url)
    api_url = (
        f"https://graph.microsoft.com/v1.0/"
        f"shares/u!{encoded_url}/driveItem/children"
    )

    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(api_url,headers=headers)

    if r.status_code != 200:
        raise Exception(f"Lỗi đọc folder:\n{r.text}")

    items = r.json().get("value", [])
    folders = []
    for item in items:
        if "folder" in item:
            folders.append({
                "name": item["name"],
                "id": item["id"],
                "webUrl": item["webUrl"]
            })
    return folders

#  ==== SCAN TREE =========================================================
def scan_folder_tree(token,current_url,result,parent_path=""):
    try:
        folders = get_folders_from_share_url(token,current_url)
        for folder in folders:
            folder_name = folder["name"]
            # ==== FULL PATH =============================================
            if parent_path:

                full_path = (
                    f"{parent_path}/"
                    f"{folder_name}"
                )
            else:
                full_path = folder_name

            # ==== SAVE DATA =============================================
            result[full_path] = {
                "name": folder_name,
                "full_path": full_path,
                "id": folder["id"],
                "webUrl": folder["webUrl"]
            }
            print("📂",full_path)

            # ==== RECURSIVE =============================================
            try:
                scan_folder_tree(token,folder["webUrl"],result,full_path)
            except Exception as e:
                print("❌ Recursive Error:",e)
    except Exception as e:
        print("❌ Scan Error:",e)

# ==== SAVE JSON =========================================================
def save_data_link_json(data):
    with open(
        DATA_LINK_FILE,"w",encoding="utf-8") as f:
        json.dump(data,f,indent=4,ensure_ascii=False)

# ==== LOAD JSON =========================================================
def load_data_link_json():
    if not os.path.exists(DATA_LINK_FILE):
        return {}
    with open(DATA_LINK_FILE,"r",encoding="utf-8"
    ) as f:
        return json.load(f)

# ==== BUILD JSON =========================================================
def build_data_link():
    token = graph_session.ensure_token()
    result = {}

    print("🔍 Đang scan OneDrive...")
    scan_folder_tree(
        token,
        BASE_URL,
        result
    )
    save_data_link_json(result)

    print("✅ Đã tạo data_link.json")

# chỉ load cache local nếu đã tồn tại
DATA_LINK = load_data_link_json()
# ==== HELPER =========================================================
def get_folder_url(folder_name):
    folder = DATA_LINK.get(folder_name)
    if not folder:
        return None
    return folder["webUrl"]

# ==== AUTO URL =========================================================
FOLDER_URLS = {}
for folder_name, folder_info in DATA_LINK.items():
    FOLDER_URLS[folder_name] = folder_info["webUrl"]
print("✅ Loaded Folder URLs:")
print(FOLDER_URLS.keys())

# ==== Lấy danh sách file từ link chia sẻ ====
def list_files_from_url(share_url):
    token = graph_session.ensure_token()
    encoded_url = base64.b64encode(share_url.encode("utf-8")).decode("utf-8")
    encoded_url = encoded_url.rstrip("=").replace("/", "_").replace("+", "-")
    url = f"https://graph.microsoft.com/v1.0/shares/u!{encoded_url}/driveItem/children"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        items = r.json().get("value", [])
        return [{"id": item["id"], "name": item["name"]} for item in items if "file" in item]
    else:
        return []

# ==== Tải file từ OneDrive (sử dụng token truyền vào) ====
# ==== DOWNLOAD FILE =========================================================
def download_file(token, share_url, file_id, filename, save_dir):
    os.makedirs(save_dir, exist_ok=True)
    cache_path = os.path.join(save_dir, filename)
    # =========================================================
    # ENCODE SHARE URL
    # =========================================================
    encoded_url = base64.b64encode(
        share_url.encode("utf-8")
    ).decode("utf-8")

    encoded_url = (
        encoded_url.rstrip("=")
        .replace("/", "_")
        .replace("+", "-")
    )
    # =========================================================
    # LẤY DOWNLOAD URL
    # =========================================================
    api_url = (
        "https://graph.microsoft.com/v1.0/"
        f"shares/u!{encoded_url}/driveItem/children"
    )

    headers = {
        "Authorization": f"Bearer {token}"
    }

    r = requests.get(api_url, headers=headers)
    if r.status_code != 200:
        print(f"❌ Không lấy được danh sách file: {r.text}")
        return None
    items = r.json().get("value", [])
    target_item = None

    for item in items:
        if item.get("id") == file_id:
            target_item = item
            break

    if not target_item:
        print(f"❌ Không tìm thấy file: {filename}")
        return None
    # =========================================================
    # DOWNLOAD URL
    # =========================================================
    download_url = target_item.get(
        "@microsoft.graph.downloadUrl"
    )
    if not download_url:
        print(f"❌ Không có downloadUrl: {filename}")
        return None
    # =========================================================
    # DOWNLOAD FILE
    # =========================================================
    r = requests.get(download_url, stream=True)
    if r.status_code == 200:
        with open(cache_path, "wb") as f:
            for chunk in r.iter_content(1024):
                if chunk:
                    f.write(chunk)
        print(f"⬇️ Downloaded: {filename}")
        return cache_path
    else:
        print(
            f"❌ Lỗi tải file {filename}: "
            f"{r.status_code}"
        )
        return None

# >> Quản lý và lưu trữ METADATA để thực hiện cập nhật và đồng bộ dữ liệu
# ==== Đường dẫn file metadata ====
METADATA_FILE = os.path.join(METADATA_DIR, "onedrive_metadata.json")
# ==== Đọc metadata đã lưu ====
def load_metadata():
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

# ==== Ghi metadata mới ====
def save_metadata(metadata):
    with open(METADATA_FILE, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=4, ensure_ascii=False)

# ==== SYNC FILES FROM ONEDRIVE =========================================================
def sync_files_from_onedrive(token, share_url, save_dir):
    """
    Đồng bộ file từ OneDrive về thư mục save_dir.
    Chức năng:
    - Lấy danh sách file từ OneDrive
    - Kiểm tra metadata local
    - Nếu file mới hơn trên cloud -> tải lại
    - Nếu chưa có -> tải về
    - Lưu metadata
    """
    # ==== TẠO THƯ MỤC NẾU CHƯA CÓ =====================================================
    os.makedirs( save_dir, exist_ok=True)
    # ==== HELPER: TÌM FILE LOCAL =====================================================
    def find_local_paths_by_name(filename):
        filename_lower = filename.lower()
        found = []
        for dirpath, dirs, files in os.walk(save_dir):
            for f in files:
                if f.lower() == filename_lower:
                    found.append(
                        os.path.join(
                            dirpath,
                            f
                        )
                    )
        return found

    # ==== ENCODE SHARE URL =====================================================
    encoded_url = base64.b64encode(
        share_url.encode("utf-8")
    ).decode("utf-8")

    encoded_url = (
        encoded_url.rstrip("=")
        .replace("/", "_")
        .replace("+", "-")
    )

    # ==== API URL =====================================================
    url = (
        "https://graph.microsoft.com/v1.0/"
        f"shares/u!{encoded_url}/driveItem/children"
    )

    headers = {
        "Authorization": f"Bearer {token}"
    }

    r = requests.get(
        url,
        headers=headers
    )

    if r.status_code != 200:
        print(
            f"❌ Không đọc được folder:\n{share_url}"
        )
        return

    items = r.json().get("value",[])
    files = [
        item for item in items
        if "file" in item
    ]

    # ==== LOAD METADATA =====================================================
    local_metadata = load_metadata()

    # ==== LOOP FILES =====================================================
    for file in files:
        file_id = file.get("id")
        file_name = file.get("name")
        last_modified = file.get(
            "lastModifiedDateTime"
        )

        # ==== PARSE REMOTE TIME =====================================================
        try:
            remote_dt = datetime.datetime.fromisoformat(
                last_modified.replace(
                    "Z",
                    "+00:00"
                )
            )
            remote_ts = remote_dt.timestamp()
        except:
            remote_ts = None

        # ==== TÌM FILE LOCAL =====================================================
        candidate_paths = []
        # metadata path
        if (
            file_id in local_metadata
            and "local_path" in local_metadata[file_id]
        ):
            candidate_paths.append(
                local_metadata[file_id]["local_path"]
            )

        # scan folder
        candidate_paths.extend(
            find_local_paths_by_name(
                file_name
            )
        )

        # remove duplicate
        seen = set()
        candidate_paths = [
            p for p in candidate_paths
            if not (
                p in seen
                or seen.add(p)
            )
        ]
        need_download = False

        # ==== CHECK FILE LOCAL =====================================================
        if not candidate_paths:
            need_download = True
        else:
            local_is_fresh = False
            for p in candidate_paths:
                try:
                    if os.path.exists(p):
                        local_ts = os.path.getmtime(p)
                        # nếu parse lỗi time
                        if remote_ts is None:
                            try:
                                os.remove(p)
                            except:
                                pass
                            need_download = True
                            break
                        # file local cũ hơn
                        if local_ts < (remote_ts - 1):
                            try:
                                os.remove(p)
                                print(
                                    f"♻️ Cập nhật file: {file_name}"
                                )
                            except Exception as e:
                                print(
                                    f"❌ Không thể xóa {p}: {e}"
                                )
                            need_download = True
                        else:
                            local_is_fresh = True
                            local_metadata[file_id] = {
                                "name": file_name,
                                "lastModifiedDateTime": last_modified,
                                "local_path": p
                            }
                            break
                    else:
                        need_download = True
                except Exception as e:
                    print(
                        f"❌ Lỗi xử lý local file:\n{e}"
                    )
                    need_download = True
            if (
                not local_is_fresh
                and not need_download
            ):
                need_download = True

        # ==== DOWNLOAD FILE =================================================
        if need_download:
            filepath = download_file(token,share_url,file_id,file_name,save_dir)
            if filepath:
                print(f"⬇️ Saved: {file_name}")
                local_metadata[file_id] = {
                    "name": file_name,
                    "lastModifiedDateTime": last_modified,
                    "local_path": filepath
                }

    # ==== SAVE METADATA =====================================================
    save_metadata(local_metadata)

# ==== Bắt đầu đồng bộ dữ liệu từ OneDrive ===================================================== 
# ==== GET SAVE DIRECTORY FROM PATH =====================================================
def get_save_dir_from_path(folder_path):
    path_upper = folder_path.upper()
    
    # Nhận diện phân hệ nguồn từ OneDrive
    is_aeongms = "AEONGMS" in path_upper
    is_maxvalu = "MAXVALU" in path_upper

    # =====================================================
    # SENSOR
    # =====================================================
    if "SENSOR" in path_upper:
        if is_aeongms: return AEONGMS_IMAGE_SENSOR_ARCHIVE_DIR
        if is_maxvalu: return MAXVALU_IMAGE_SENSOR_ARCHIVE_DIR
        return None

    # =====================================================
    # GATEWAY
    # =====================================================
    elif "GATEWAY" in path_upper:
        if is_aeongms: return AEONGMS_IMAGE_GATEWAY_ARCHIVE_DIR
        if is_maxvalu: return MAXVALU_IMAGE_GATEWAY_ARCHIVE_DIR
        return None

    # =====================================================
    # LAYOUT
    # =====================================================
    elif "LAYOUT" in path_upper:
        if is_aeongms: return AEONGMS_IMAGE_LAYOUT_ARCHIVE_DIR
        if is_maxvalu: return MAXVALU_IMAGE_LAYOUT_ARCHIVE_DIR
        return None

    # =====================================================
    # ALARM POINT / ALARM POINTS
    # =====================================================
    elif "ALARM_POINT" in path_upper or "ALARM_POINTS" in path_upper:
        if is_aeongms: return AEONGMS_IMAGE_AL_ARCHIVE_DIR
        if is_maxvalu: return MAXVALU_IMAGE_AL_ARCHIVE_DIR
        return None

    # =====================================================
    # REPORT FORM - AEON GMS
    # =====================================================
    elif (
        "REPORT_FORM/AEONGMS" in path_upper
        or "REPORT FORM/AEONGMS" in path_upper
        or "REPORT_FORM\\AEONGMS" in path_upper
    ):
        return AEONGMS_DIR

    # =====================================================
    # REPORT FORM - MAXVALU
    # =====================================================
    elif (
        "REPORT_FORM/MAXVALU" in path_upper
        or "REPORT FORM/MAXVALU" in path_upper
        or "REPORT_FORM\\MAXVALU" in path_upper
    ):
        return AEONMAXVALU_DIR

    # =====================================================
    # REPORT FORM CHUNG
    # =====================================================
    elif (
        "REPORT_FORM" in path_upper
        or "REPORT FORM" in path_upper
        or "HOTLINE_AND_CONFIRM_FORM" in path_upper
    ):
        return REPORT_FORM_DIR

    # =====================================================
    # DOCUMENTARY
    # =====================================================
    elif "DOCUMENTARY" in path_upper:
        return DOCUMENTARY_ARCHIVE_DIR

    # =====================================================
    # NOTE
    # =====================================================
    elif "NOTE" in path_upper:
        return NOTE_ARCHIVE_DIR

    return None

# ==== AUTO SYNC FROM data_link.json =========================================================
def auto_sync_all_onedrive():
    token = graph_session.ensure_token()
    data_links = load_data_link_json()
    if not data_links:
        print("❌ Không tìm thấy data_link.json")
        return
    print("🔄 Bắt đầu đồng bộ dữ liệu...")

    total = 0
    success = 0

    for folder_name, folder_info in data_links.items():
        try:
            share_url = folder_info.get("webUrl")
            if not share_url:
                continue
            # ==== AUTO DETECT SAVE DIR =============================================
            folder_path = folder_info.get("full_path","")
            save_dir = get_save_dir_from_path(folder_path)
            if not save_dir:
                print(
                    f"⚠️ Không xác định được thư mục lưu: {folder_name}"
                )
                continue
            print(f"📂 Sync: {folder_name}")
            print(f"💾 Save to: {save_dir}")
            # ==== SYNC =============================================
            sync_files_from_onedrive(token,share_url,save_dir)
            success += 1
        except Exception as e:
            print(
                f"❌ Lỗi sync {folder_name}:",e
            )
        total += 1
    print(
        f"✅ Đồng bộ hoàn tất "
        f"({success}/{total})"
    )

# === Khu vực tạo Frame để lưu trữ các thành phần ===============================================================================================
# Tạo cửa sổ chính
root = tk.Tk()
root.title("RMC Report Assistant")
root.geometry("1366x768")

# Frame chính
main_frame = tk.Frame(root)
main_frame.pack(expand=True, pady=40, padx=20)

# Frame con chứa văn bản và các nút bên phải
content_frame = tk.Frame(main_frame)
content_frame.pack()

# === frame chứa nút contact, status và note bên trái ===
left_button_frame = tk.Frame(content_frame)
left_button_frame.pack(side="left", fill="y", padx=10, pady=10)

# === Text để hiển thị văn bản ===
output_text = tk.Text(content_frame, font=("Arial", 13), width=60, height=17, wrap="word")
output_text.pack(side='left', pady=(10, 0), padx=10)
output_text.config(state='disabled')

# === Frame chứa các item xuất hiện khi chọn danh sách ===
item_frame = tk.Frame(content_frame)
item_frame.pack(side='left', padx=10)

# ==== NÚT COPY ====
copy_frame = tk.Frame(main_frame)
copy_frame.pack(fill='x', pady=(10, 0), padx=20)

# Nhóm bên phải: Catch, Clock, Continue
right_controls = tk.Frame(copy_frame)
right_controls.pack(side="right")

# Nhóm bên trái: Copy và Clear
left_controls = tk.Frame(copy_frame)
left_controls.pack(side="left")

# ==== Frame chứa các ô tô màu ====
box_frame = tk.Frame(main_frame)
box_frame.pack(pady=(10, 0))

# === Khu vực tạo và cấu hình chức năng ===========================================================================================
# ============== Chức năng tô màu ô tiến trình ==================
boxes = [] # Danh sách các ô vuông
box_colors = ["white"] * 6
hint_label = None # Label để hiển thị gợi ý
box_filled = [False] * 6
first_box_filled = False

# ==== Tạo 6 ô trắng trong box_frame ====
for i in range(6):
    lbl = tk.Label(box_frame, width=10, height=1, bg="white", relief="solid", borderwidth=2)
    lbl.grid(row=0, column=i, padx=5)
    boxes.append(lbl)

# =========================================================
# KHUNG CHỨA HINT CỐ ĐỊNH KÍCH THƯỚC
# =========================================================
# Định nghĩa chiều rộng (width) và chiều cao (height) cố định cho khung gợi ý
hint_container = tk.Frame(main_frame, width=800, height=30, bg="SystemButtonFace")
hint_container.pack(pady=(10, 20))
hint_container.pack_propagate(False) # KHÓA CỨNG KÍCH THƯỚC: Ép khung không giãn theo chữ

# Đưa hint_label vào trong hint_container (bỏ bớt phần wraplength tĩnh)
hint_label = tk.Label(
    hint_container, 
    text="Quy trình xử lý sự cố đang đợi", 
    justify="left", 
    font=("Arial", 10), 
    fg="black"
)
# Sử dụng fill="both" và expand=True để chữ tự căn đều trong khung cố định
hint_label.pack(fill="both", expand=True)

# ==== Hàm xử lý tô màu ====
def on_category_click():
    global box_colors
    # Đếm số ô đã được tô xanh
    green_count = box_colors.count("green")
    # Gợi ý tương ứng từng bước
    if green_count == 1:
        update_hint("Đã ghi nhận sự cố, tiến hành báo cáo lên group chung và tiếp tục theo dõi sự cố đang diễn ra. Nếu trong vòng 5 phút, không có thông báo gì từ phía bên Site đang xảy ra lỗi lên group chung. Lập tức liên hệ vói Site theo danh sách đã cho dựa vào mức độ ưu tiên (Bấm xác nhận nếu như thông tin đã được cập nhật lên group từ bên Site). Sau khi đã liên hệ, cập nhật thông tin liên hệ lên Group chung thông qua biểu mẫu trong mục Contact.")
    elif green_count == 2:
        update_hint("Tiếp tục theo dõi và cập nhật sự cố liên tục. Nếu như sau một khoảng thời gian không nhận được thông tin gì từ phía bên Site kể từ thời điểm đã liên hệ với Site (1 - 2 tiếng) và đã thông tin lên group chung (). Tiến hành liên hệ lại với Site để xác minh tình trạng kiếm tra thiết bị và nguyên nhân (nếu có). Tiến hành cập nhập lại tình hình thiết bị lên nhóm group chung về tình hình khắc phục trình trạng hiện tại của thiết bị gây lỗi.")
    elif green_count == 3:
        update_hint("Nếu sự cố sau 1 tiếng cho đến 2 tiếng vẫn chưa được sự xử lí và cũng chưa được cập nhật lên group chung. Tiến hành liên hệ lại với số điện thoại ưu tiên để xác nhận lại sự cố, sau đó báo cáo lại tình hiên lên group chung (Bấm 'Xác nhận' nếu sự cố đã được giải quyết trước thời điểm này).")
    elif green_count == 4:
        update_hint("Khi sự cố đã được giải quyết, báo cáo lên group chung để khách hàng và các bộ phận liên quan nắm thông tin (Bấm 'Xác nhận' nếu có trường hợp ngoại lệ xảy ra).")
    elif green_count == 5:
        update_hint("Cập nhật lên bảng Alarm List.")
    elif green_count == 6:
        update_hint("Toàn bộ các bước trong quy trình đã được hoàn tất, làm tốt lắm!")
        root.after(5000,reset_after_delay)

def reset_after_delay():
    if not root.winfo_exists():
        return

    for i in range(6):
        box_colors[i] = "white"
        box_filled[i] = False
        if boxes[i].winfo_exists():
            boxes[i].config(bg="white")
    global first_box_filled
    first_box_filled = False
    update_hint(
        "Quy trình xử lý sự cố đang đợi"
    )
# =========================================================
# CLEANUP ON EXIT
# =========================================================
# =========================================================
# CLEANUP ON EXIT
# =========================================================
def on_closing():
    """Cancel scheduled after jobs and exit cleanly."""
    global clock_after_id, countdown_job, schedule_after_id, reset_after_id, schedule_running, scroll_after_id
    
    # stop schedule loop
    try:
        schedule_running = False
    except Exception:
        pass
        
    # cancel known after jobs
    try:
        if clock_after_id:
            root.after_cancel(clock_after_id)
    except Exception:
        pass
    try:
        if countdown_job:
            root.after_cancel(countdown_job)
    except Exception:
        pass
    try:
        if schedule_after_id:
            root.after_cancel(schedule_after_id)
    except Exception:
        pass
    try:
        if reset_after_id:
            root.after_cancel(reset_after_id)
    except Exception:
        pass
        
    # Thêm phần dọn dẹp tiến trình cuộn chữ
    try:
        if scroll_after_id:
            root.after_cancel(scroll_after_id)
    except Exception:
        pass
        
    try:
        root.destroy()
    except Exception:
        try:
            root.quit()
        except Exception:
            pass

# =========================================================
# XỬ LÝ CHỮ CHẠY (MARQUEE) CHO HINT
# =========================================================
scroll_after_id = None
full_hint_text = ""
# Cài đặt số ký tự tối đa hiển thị trên khung (Bạn có thể tăng/giảm tùy kích thước khung)
DISPLAY_LEN = 80 

def update_hint(text):
    global scroll_after_id, full_hint_text
    
    # Nếu đang có hiệu ứng chạy chữ cũ thì phải hủy nó đi trước khi hiển thị chữ mới
    if scroll_after_id is not None:
        try:
            root.after_cancel(scroll_after_id)
        except Exception:
            pass
        scroll_after_id = None
        
    if hint_label:
        # Nếu chữ ngắn gọn, không cần chạy chữ
        if len(text) <= DISPLAY_LEN:
            hint_label.config(text=text)
        else:
            # Nếu chữ dài, thêm một khoảng trắng ở cuối để tạo sự ngắt quãng khi lặp lại
            full_hint_text = text + " " * 30 
            scroll_text()

def scroll_text():
    global scroll_after_id, full_hint_text
    
    if hint_label and hint_label.winfo_exists():
        # Lấy một đoạn chuỗi vừa đủ độ dài khung để hiển thị
        display_text = full_hint_text[:DISPLAY_LEN]
        hint_label.config(text=display_text)
        
        # Đưa ký tự đầu tiên xuống cuối cùng để text dịch chuyển sang trái
        full_hint_text = full_hint_text[1:] + full_hint_text[0]
        
        # Chạy lại hàm này sau 120ms (Giảm số này chữ sẽ chạy nhanh hơn, tăng thì chậm lại)
        scroll_after_id = root.after(120, scroll_text)

# ==== Chức năng cho nút copy và nút clear văn bản đang hiển thị trên text box ====
def copy_text_to_clipboard():
    text = output_text.get("1.0", "end-1c")
    pyperclip.copy(text)

def clear_text_output():
    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.config(state='disabled')

# == Chức năng bắt và tiếp tục đồng hồ thời gian thực của hệ thống ==
# =========================================================
# CLOCK GLOBAL
# =========================================================
clock_after_id = None
countdown_job = None
clock_after_id = None
is_running = True
clock_running = False

# =========================================================
# PAUSE CLOCK
# =========================================================
def catch_clock():
    global is_running
    is_running = False

# =========================================================
# CONTINUE CLOCK
# =========================================================
def continue_clock():
    global is_running
    is_running = True
    start_clock()

# == Chức năng lấy và hiển thị thời gian hiện tại trên hệ thống lên văn bản ==
# =========================================================
# TIMER GLOBAL
# =========================================================
countdown_job = None
time_left = 300
reset_after_id = None
# =========================================================
# UPDATE TIMER
# =========================================================
def update_timer():
    global time_left
    global countdown_job
    try:
        if not root.winfo_exists():
            return
        if not timer_label.winfo_exists():
            return
        if time_left <= 0:
            timer_label.config(
                text="⏰ Hết thời gian"
            )
            countdown_job = None
            return
        minutes = time_left // 60
        seconds = time_left % 60
        timer_label.config(
            text=f"Running... {minutes:02}:{seconds:02}"
        )
        time_left -= 1
        # lưu ID
        countdown_job = root.after(
            1000,
            update_timer
        )
    except:
        countdown_job = None
# =========================================================
# UPDATE CLOCK
# =========================================================
def update_clock():
    global clock_after_id
    global clock_running
    try:
        if not root.winfo_exists():
            clock_running = False
            return

        if not clock_label.winfo_exists():
            clock_running = False
            return

        current_time = datetime.datetime.now().strftime(
            "%H:%M:%S"
        )

        clock_label.config(
            text=current_time
        )

        clock_after_id = root.after(
            1000,
            update_clock
        )
    except tk.TclError:
        clock_running = False
        clock_after_id = None
# =========================================================
# START & STOP SYSTEM CLOCK
# =========================================================
def start_clock():
    global clock_running
    if clock_running:
        return
    clock_running = True
    update_clock()
def stop_clock():
    global clock_running
    global clock_after_id
    clock_running = False
    if clock_after_id:
        try:
            root.after_cancel(clock_after_id)
        except:
            pass
        clock_after_id = None
# == Chức năng bắt đầu và reset đồng hồ đếm ngược ==
# =========================================================
# START TIMER
# =========================================================
def start_timer():
    global time_left
    global countdown_job
    try:
        # =========================================
        # HỦY TIMER CŨ
        # =========================================
        if countdown_job:
            root.after_cancel(
                countdown_job
            )
            countdown_job = None
    except:
        pass
    # =============================================
    # RESET
    # =============================================
    time_left = 300
    # =============================================
    # START
    # =============================================
    update_timer()
# =========================================================
# RESET TIMER
# =========================================================
def reset_timer():
    global time_left
    global countdown_job
    try:
        if countdown_job:
            root.after_cancel(
                countdown_job
            )
    except:
        pass
    countdown_job = None
    time_left = 300
    try:
        if timer_label.winfo_exists():
            timer_label.config(
                text="⏳Waiting Countdown⏳"
            )
    except:
        pass

# ==== thêm đồng hồ đếm ngược ==== 
timer_frame = tk.Frame(main_frame)
timer_frame.pack(pady=(10, 0))

timer_label = tk.Label(timer_frame, text="⏳Waiting Countdown⏳", font=("Arial", 12, "bold"), fg="blue")
timer_label.pack()

countdown_job = None
time_left = 300  # 5 phút = 300 giây

# ==== Đăng nhập, tải file và xử lý OneDrive bằng Azure ===========================================================================
# ==== Hàm đăng nhập Azure AD ====
def authenticate():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r").read())

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()

    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if "user_code" not in flow:
            raise Exception("Không khởi tạo được Device Flow")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())
        return result["access_token"]
    else:
        raise Exception("Đăng nhập thất bại: " + str(result))
# Đăng nhập Azure
access_token = authenticate()
# ==== Lấy danh sách file từ link chia sẻ ====
def list_files_from_url(token, share_url):
    encoded_url = base64.b64encode(share_url.encode("utf-8")).decode("utf-8")
    encoded_url = encoded_url.rstrip("=").replace("/", "_").replace("+", "-")
    url = f"https://graph.microsoft.com/v1.0/shares/u!{encoded_url}/driveItem/children"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        items = r.json().get("value", [])
        return [{"id": item["id"], "name": item["name"]} for item in items if "file" in item]
    else:
        return []
# === Hiển thị file văn bản từ OneDrive ===========================================================================
# ==== LẤY DANH SÁCH FILE ONE DRIVE THEO TÊN ====
# =========================================================
# BUILD DEVICE MAPPING FROM LOCAL FOLDER
# =========================================================
def build_device_mapping_from_local(report_dir):
    mapping = {}
    IGNORE_FILES = {"CONTACT_FORM","CONFIRM_FORM","NOTIFICATION_FORM"}
    # ==========================================
    # QUÉT TOÀN BỘ THƯ MỤC CON
    # ==========================================
    for root, dirs, files in os.walk(report_dir):
        for filename in files:
            file_path = os.path.join(
                root,
                filename
            )
            # ==================================
            # REMOVE EXTENSION
            # ==================================
            name_without_ext = os.path.splitext(
                filename
            )[0].strip()
            upper_name = name_without_ext.upper()
            # ==================================
            # IGNORE
            # ==================================
            if upper_name in IGNORE_FILES:
                continue
            # ==================================
            # FORMAT:
            # NVL_FR&FC
            # ==================================
            if "_" not in name_without_ext:
                continue
            parts = name_without_ext.split(
                "_",
                1
            )
            if len(parts) < 2:
                continue
            area = parts[0].strip().upper()
            device = parts[1].strip().upper()
            if not area or not device:
                continue
            if area not in mapping:
                mapping[area] = {}
            # ==================================
            # SAVE PATH
            # ==================================
            mapping[area][device] = file_path
    return mapping
def refresh_report_mapping():

    global ALL_REPORT_MAPPINGS
    global REPORT_FORM_MAPPING

    ALL_REPORT_MAPPINGS["AEONGMS"] = \
        build_device_mapping_from_local(
            AEONGMS_DIR
        )

    ALL_REPORT_MAPPINGS["MAXVALU"] = \
        build_device_mapping_from_local(
            AEONMAXVALU_DIR
        )

    REPORT_FORM_MAPPING = \
        ALL_REPORT_MAPPINGS[CURRENT_SOURCE]

    print("================================")
    print("REPORT_FORM_MAPPING REFRESHED")
    print(
        json.dumps(
            REPORT_FORM_MAPPING,
            indent=4,
            ensure_ascii=False
        )
    )
    print("================================")
# ==== CHỨC NĂNG HIỂN THỊ VĂN BẢN VÀ THỜI GIAN====
# =========================================================
# SHOW TEXT FROM LOCAL FILE
# =========================================================
def show_text_from_local(
    file_path,
    is_no_error=False,
    start_timer_flag=True
):
    try:
        with open(file_path,'r',encoding='utf-8',errors='ignore'
        ) as f:
            lines = f.readlines()
        # =============================================
        # NO ERROR
        # =============================================
        if is_no_error:
            yesterday = (
                datetime.datetime.now()
                - datetime.timedelta(days=1)
            )
            timestamp = (
                yesterday.strftime(
                    "Trong ngày: %d-%m-%Y "
                ) + '\n'
            )
            lines = [
                timestamp
                if '[no_error_time]' in line
                else line
                for line in lines
            ]
        # =============================================
        # NORMAL
        # =============================================
        else:
            delayed_time = (
                datetime.datetime.now()
                - datetime.timedelta(minutes=1)
            )
            current_time = (
                delayed_time.strftime(
                    "+ Thời gian: %H:%M:%S %d-%m-%Y "
                ) + '\n'
            )
            lines = [
                current_time
                if '[time]' in line
                else line
                for line in lines
            ]
        content = ''.join(lines)
    except Exception as e:
        content = f"Không thể mở file: {e}"
    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.insert(
        tk.END,
        content
    )
    output_text.config(state='disabled')
    if start_timer_flag:
        start_timer()

REPORT_FORM_MAPPING = build_device_mapping_from_local(
    REPORT_FORM_DIR
)

print("✅ REPORT FORM MAPPING:")
print(json.dumps(
    REPORT_FORM_MAPPING,indent=4,
    ensure_ascii=False
))

# ==== TẠO GIAO DIỆN DANH SÁCH ====
active_parent_button = None
active_child_button = None

def set_active_parent_button(btn):
    global active_parent_button
    # Reset nút cha cũ
    if active_parent_button and active_parent_button != btn:
        active_parent_button.config(bg="SystemButtonFace", fg="black")
    # Đổi màu nút cha mới
    btn.config(bg="green", fg="white")
    active_parent_button = btn
def set_active_child_button(btn):
    global active_child_button
    # Reset nút con cũ
    if active_child_button and active_child_button != btn:
        active_child_button.config(bg="SystemButtonFace", fg="black")
    # Đổi màu nút con mới
    btn.config(bg="blue", fg="white")
    active_child_button = btn
def create_list_block(parent, list_name, items, toggle_function, state):
    block_frame = tk.Frame(parent)
    # Đổi anchor='w' thành fill='x' để khung tự động dãn ngang
    block_frame.pack(pady=10, fill='x', padx=5) # Thêm padx=5 để nút không sát rạt vào lề

    # Bỏ thuộc tính width=10 để nút linh hoạt kích thước
    list_button = tk.Button(
        block_frame,
        text=list_name,
        font=("Arial", 14),
        command=lambda: [set_active_parent_button(list_button), toggle_function(state)]
    )
    # Đổi anchor='w' thành fill='x' để bản thân nút tự dãn lấp đầy block_frame
    list_button.pack(fill='x')
    state["button"] = list_button

# ==== HÀM BẬT TẮT DANH SÁCH ====
# =========================================================
# KHUNG LƯU HAI NÚT BẤM
# =========================================================
source_frame = tk.Frame(content_frame,bg="white")
source_frame.pack(fill="x",pady=(5, 10))
# =========================================================
# HÀM ĐỔI NGUỒN
# =========================================================
# ==== Chức năng đổi màu nút nguồn ===
active_source_button = None
def set_active_source_button(btn):
    global active_source_button
    try:
        if (
            active_source_button
            and active_source_button.winfo_exists()
            and active_source_button != btn
        ):

            active_source_button.config(bg="SystemButtonFace",fg="black"
            )
    except:
        pass
    active_source_button = btn
    try:
        active_source_button.config(bg="#2196F3",fg="white")
    except:
        pass
# ==== Chức năng nút nguồn ===
def switch_source(source_name, button):
    global CURRENT_SOURCE
    global REPORT_FORM_MAPPING
    global current_open_area

    CURRENT_SOURCE = source_name
    REPORT_FORM_MAPPING = \
        ALL_REPORT_MAPPINGS[source_name]
    current_open_area = None
    # Đổi màu nút
    set_active_source_button(button)
    create_area_buttons()
    # Xóa danh sách thiết bị cũ
    for widget in device_scrollable_frame.winfo_children():
        widget.destroy()
    # Reset tìm kiếm
    search_parent_var.set("")
    search_device_var.set("")

btn_aeongms = tk.Button(source_frame,text="AEONGMS",font=("Arial", 11, "bold"),width=15)
btn_aeongms.config(command=lambda:switch_source("AEONGMS",btn_aeongms))
btn_aeongms.pack(side="left",padx=5)

btn_maxvalu = tk.Button(source_frame,text="MAXVALU",font=("Arial", 11, "bold"),width=15)
btn_maxvalu.config(command=lambda:switch_source("MAXVALU",btn_maxvalu))
btn_maxvalu.pack(side="left",padx=5)

root.after(100,lambda: switch_source("AEONGMS",btn_aeongms))
# =========================================================
# MAIN CONTAINER
# =========================================================
main_container = tk.Frame(content_frame,bg="white")
main_container.pack(fill="both",expand=True,padx=10,pady=10)
# =========================================================
# AREA CONTAINER
# =========================================================
area_container = tk.Frame(main_container,width=180,bg="white")
area_container.pack(side="left",fill="y")
area_container.pack_propagate(False)
# =========================================================
# SITE SEARCH FRAME
# =========================================================
site_search_frame = tk.Frame(area_container,bg="white")
site_search_frame.pack(fill="x",padx=5,pady=(5, 0))
# =========================================================
# DEVICE CONTAINER
# =========================================================
# 1. Thêm width cố định (ví dụ 220 hoặc 250 tùy độ dài tên thiết bị)
device_container = tk.Frame(main_container, width=220, bg="white")

# 2. Xóa expand=True và đổi fill="both" thành fill="y" để khung chỉ kéo dài theo chiều dọc
device_container.pack(side="left", fill="y", padx=(10, 0))

# 3. Thêm dòng này để khóa cứng chiều rộng, ép thanh cuộn phải bám sát vào lề của 220px
device_container.pack_propagate(False)

# =========================================================
# DEVICE SEARCH FRAME
# =========================================================
device_search_frame = tk.Frame(device_container,bg="white")
device_search_frame.pack(fill="x",padx=5,pady=(5, 0))
# =========================================================
# SEARCH AREA
# =========================================================
search_parent_var = tk.StringVar()
search_parent_entry = tk.Entry(site_search_frame,textvariable=search_parent_var,font=("Arial", 10))
search_parent_entry.pack(fill="x",pady=5)
# =========================================================
# SEARCH DEVICE
# =========================================================
search_device_var = tk.StringVar()
search_device_entry = tk.Entry(device_search_frame,textvariable=search_device_var,font=("Arial", 10))
search_device_entry.pack(fill="x",pady=5)
# =========================================================
# AREA CANVAS
# =========================================================
area_canvas = tk.Canvas(area_container,bg="white",highlightthickness=0)
area_scrollbar = tk.Scrollbar(area_container,orient="vertical",command=area_canvas.yview)
area_scrollable_frame = tk.Frame(area_canvas,bg="white")
# =========================================================
# AREA SCROLLBAR
# =========================================================
area_scrollbar = tk.Scrollbar(area_container,orient="vertical",command=area_canvas.yview)
area_canvas.configure(yscrollcommand=area_scrollbar.set)
area_scrollbar.pack(side="right",fill="y")

# =========================================================
# UPDATE AREA SCROLL
# =========================================================
def on_area_configure(event):
    area_canvas.configure(
        scrollregion=area_canvas.bbox("all")
    )
area_scrollable_frame.bind("<Configure>",on_area_configure)

# =========================================================
# CREATE AREA WINDOW
# =========================================================
# Lưu lại ID của window bên trong canvas
area_window_id = area_canvas.create_window((0, 0), window=area_scrollable_frame, anchor="nw")
area_canvas.configure(yscrollcommand=area_scrollbar.set)

# Thứ tự pack: Scrollbar trước, Canvas sau
area_scrollbar.pack(side="right", fill="y")
area_canvas.pack(side="left", fill="both", expand=True)

# =========================================================
# UPDATE AREA SCROLL & AUTO WIDTH
# =========================================================
def on_area_configure(event):
    area_canvas.configure(scrollregion=area_canvas.bbox("all"))

def on_area_canvas_configure(event):
    # Ép khung cuộn khu vực luôn rộng bằng Canvas chứa nó
    area_canvas.itemconfig(area_window_id, width=event.width)

area_scrollable_frame.bind("<Configure>", on_area_configure)
area_canvas.bind("<Configure>", on_area_canvas_configure)

# =========================================================
# DEVICE CANVAS
# =========================================================
device_canvas = tk.Canvas(device_container,bg="white",highlightthickness=0)
device_scrollbar = tk.Scrollbar(device_container,orient="vertical",command=device_canvas.yview)
device_scrollable_frame = tk.Frame(device_canvas,bg="white")
# =========================================================
# GẮN FRAME VÀO CANVAS
# =========================================================
# Lưu lại ID của window bên trong canvas để kiểm soát chiều rộng
device_window_id = device_canvas.create_window((0, 0), window=device_scrollable_frame, anchor="nw")
device_canvas.configure(yscrollcommand=device_scrollbar.set)
# =========================================================
# THỨ TỰ PACK CHUẨN (RẤT QUAN TRỌNG)
# =========================================================
# 1. BẮT BUỘC phải pack Scrollbar trước để nó luôn bám sát lề phải
device_scrollbar.pack(side="right", fill="y")
# 2. Pack Canvas sau để nó tự động điền vào phần không gian còn lại
device_canvas.pack(side="left", fill="both", expand=True)
# =========================================================
# XỬ LÝ SỰ KIỆN CUỘN & CHỐNG TRÀN CHIỀU NGANG
# =========================================================
def on_device_configure(event):
    # Cập nhật lại giới hạn cuộn dọc khi thêm/bớt nút thiết bị
    device_canvas.configure(scrollregion=device_canvas.bbox("all"))

def on_canvas_configure(event):
    # Ép chiều rộng của frame chứa nút bấm bằng đúng chiều rộng của Canvas.
    # Việc này ngăn tên thiết bị quá dài đâm xuyên qua lề phải làm hỏng thanh cuộn.
    device_canvas.itemconfig(device_window_id, width=event.width)

# Lắng nghe sự thay đổi kích thước
device_scrollable_frame.bind("<Configure>", on_device_configure)
device_canvas.bind("<Configure>", on_canvas_configure)

# =========================================================
# CREATE DEVICE WINDOW
# =========================================================
device_canvas.create_window((0, 0),window=device_scrollable_frame,anchor="nw")
device_canvas.configure(yscrollcommand=device_scrollbar.set)
device_canvas.pack(side="left",fill="both",expand=True)
device_scrollbar.pack(side="right",fill="y")

# =========================================================
# GLOBAL STATES
# =========================================================
area_states = {}
current_open_area = None
active_parent_button = None
active_child_button = None

# đang mở khu vực nào
current_open_area = None
parent_items = []  # list of (block_frame, button) tuples for area buttons
device_items = []  # list of (block_frame, button) tuples for device buttons

# ==== THANH TÌM KIẾM CHO DANH SÁCH CHA (area buttons) ====
# =========================================================
# HÀM XỬ LÝ LỌC VÀ ĐẨY NÚT KHU VỰC LÊN ĐẦU
# =========================================================
def filter_parent_buttons(event=None):
    # LẤY TRỰC TIẾP TỪ Ô ENTRY (Thay vì dùng search_parent_var.get())
    # Cách này giúp đảm bảo lấy đúng text hiển thị, không lo lỗi đồng bộ StringVar
    keyword = search_parent_entry.get().strip().lower()
    matches = []
    non_matches = []

    # Phân loại nút trùng khớp và không trùng khớp
    for block, btn in parent_items:
        text = btn.cget("text").lower()
        if keyword == "" or keyword in text:
            matches.append((block, btn))
        else:
            non_matches.append((block, btn))

    # BƯỚC 1: Xóa toàn bộ các khối khỏi layout để chuẩn bị xếp lại thứ tự
    for block, btn in parent_items:
        block.pack_forget()

    # BƯỚC 2: Đẩy các nút có từ khóa liên quan lên ĐẦU TIÊN
    for block, btn in matches:
        block.pack(
            fill="x",
            padx=5,
            pady=4
        )

    # BƯỚC 3: Các nút KHÔNG liên quan sẽ bị đẩy xuống PHÍA DƯỚI CÙNG
    for block, btn in non_matches:
        block.pack(
            fill="x",
            padx=5,
            pady=4
        )

    # Cập nhật lại vùng cuộn Canvas sau khi thay đổi vị trí các nút
    area_scrollable_frame.update_idletasks()
    area_canvas.configure(
        scrollregion=area_canvas.bbox("all")
    )
search_parent_entry.bind("<KeyRelease>", filter_parent_buttons)

def filter_device_buttons(event=None):
    # LẤY TRỰC TIẾP TỪ Ô ENTRY THIẾT BỊ
    keyword = search_device_entry.get().strip().lower()
    matches = []
    non_matches = []

    # Phân loại nút trùng khớp và không trùng khớp
    for block, btn in device_items:
        text = btn.cget("text").lower()
        if keyword == "" or keyword in text:
            matches.append((block, btn))
        else:
            non_matches.append((block, btn))

    # BƯỚC 1: Xóa toàn bộ các khối thiết bị khỏi layout
    for block, btn in device_items:
        block.pack_forget()

    # BƯỚC 2: Đẩy các thiết bị có từ khóa liên quan lên ĐẦU TIÊN
    for block, btn in matches:
        block.pack(
            fill="x",
            padx=5,
            pady=3
        )

    # BƯỚC 3: Các thiết bị KHÔNG liên quan đẩy xuống PHÍA DƯỚI
    for block, btn in non_matches:
        block.pack(
            fill="x",
            padx=5,
            pady=3
        )

    # LƯU Ý SỬA TẠI ĐÂY: Cập nhật lại vùng cuộn Canvas của THIẾT BỊ (Đã sửa từ area sang device)
    device_scrollable_frame.update_idletasks()
    device_canvas.configure(
        scrollregion=device_canvas.bbox("all")
    )
search_device_entry.bind("<KeyRelease>", filter_device_buttons)
# =========================================================
# SET ACTIVE PARENT BUTTON
# =========================================================
active_parent_button = None
def set_active_parent_button(btn):
    global active_parent_button
    # =====================================================
    # RESET BUTTON CŨ
    # =====================================================
    try:
        if (
            active_parent_button
            and active_parent_button.winfo_exists()
            and active_parent_button != btn
        ):
            active_parent_button.config(
                bg="SystemButtonFace",
                fg="black"
            )
    except:
        pass
    # =====================================================
    # SET BUTTON MỚI
    # =====================================================
    active_parent_button = btn
    try:
        if (
            active_parent_button
            and active_parent_button.winfo_exists()
        ):

            active_parent_button.config(
                bg="lightblue",
                fg="black"
            )
    except:
        pass
# =========================================================
# SET ACTIVE CHILD BUTTON
# =========================================================
active_child_button = None
def set_active_child_button(btn):
    global active_child_button
    # =====================================================
    # RESET BUTTON CŨ
    # =====================================================
    try:
        if (
            active_child_button
            and active_child_button.winfo_exists()
            and active_child_button != btn
        ):
            active_child_button.config(
                bg="SystemButtonFace",
                fg="black"
            )
    except:
        pass
    # =====================================================
    # SET BUTTON MỚI
    # =====================================================
    active_child_button = btn
    try:
        if (
            active_child_button
            and active_child_button.winfo_exists()
        ):
            active_child_button.config(
                bg="#4CAF50",
                fg="white"
            )
    except:
        pass
# =========================================================
# TOGGLE AREA
# =========================================================
def toggle_area(area_name):
    global current_open_area
    # =====================================================
    # CLOSE OLD AREA
    # =====================================================
    if (
        current_open_area
        and current_open_area != area_name
    ):
        old_state = area_states[current_open_area]
        hide_sub_buttons(old_state)
        old_state["visible"] = False
    # =====================================================
    # CURRENT STATE
    # =====================================================
    state = area_states[area_name]
    # =====================================================
    # CLOSE CURRENT
    # =====================================================
    if state["visible"]:
        hide_sub_buttons(state)
        state["visible"] = False
        current_open_area = None
        return
    # =====================================================
    # OPEN CURRENT
    # =====================================================
    show_sub_buttons(
        area_name,
        state,
        auto_select_first=True
    )
    state["visible"] = True
    current_open_area = area_name
    set_active_parent_button(
        state["button"]
    )
# =========================================================
# HIDE SUB BUTTONS
# =========================================================
def hide_sub_buttons(state):
    global active_child_button
    for btn in state["buttons"]:
        try:
            btn.destroy()
        except:
            pass
    state["buttons"].clear()
# =========================================================
# SHOW SUB BUTTONS
# =========================================================
def show_sub_buttons(area_name, state, auto_select_first=False):
    # CẢI TIẾN: Reset ô tìm kiếm thiết bị về rỗng mỗi khi người dùng đổi Khu vực
    search_device_entry.delete(0, tk.END)
    # Xóa giao diện thiết bị cũ
    for widget in device_scrollable_frame.winfo_children():
        widget.destroy()
    try:
        device_canvas.yview_moveto(0)
    except Exception:
        pass
    state["buttons"].clear()
    # CỐT LÕI: Làm sạch danh sách thiết bị cũ để nạp danh sách thiết bị mới
    device_items.clear() 
    devices = REPORT_FORM_MAPPING.get(area_name, {})
    first_child_btn = None
    for idx, (device_name, file_path) in enumerate(devices.items()):
        # Tạo block_frame cho từng thiết bị để hỗ trợ việc pack/pack_forget khi lọc
        block_frame = tk.Frame(device_scrollable_frame, bg="white")
        btn = tk.Button(
            block_frame,
            text=device_name,
            font=("Arial", 11),
            height=2
        )
        # Thiết lập tập lệnh xử lý sự kiện click nút
        if "NO_ERROR" in device_name.upper():
            cmd = lambda p=file_path, b=btn: [
                set_active_child_button(b),
                handle_first_box_fill(),
                show_text_from_local(p, is_no_error=True, start_timer_flag=False)
            ]
        else:
            cmd = lambda p=file_path, b=btn: [
                set_active_child_button(b),
                handle_first_box_fill(),
                show_text_from_local(p, start_timer_flag=True)
            ]
        btn.config(command=cmd)
        btn.pack(fill="x", padx=5, pady=3)
        block_frame.pack(fill="x", padx=5, pady=3)
        # Lưu trữ trạng thái vào hệ thống
        state["buttons"].append(btn)
        device_items.append((block_frame, btn)) # Đẩy vào mảng tracking lọc dữ liệu
        if idx == 0:
            first_child_btn = btn
    if auto_select_first and first_child_btn:
        first_child_btn.invoke()
    # Cập nhật thanh cuộn thiết bị ban đầu
    device_scrollable_frame.update_idletasks()
    device_canvas.configure(scrollregion=device_canvas.bbox("all"))
# =========================================================
# AUTO CREATE AREA BUTTONS
# =========================================================
def create_area_buttons():
    # =============================================
    # CLEAR BUTTON CŨ
    # =============================================
    for widget in area_scrollable_frame.winfo_children():
        widget.destroy()
    # =============================================
    # CREATE BUTTON BLOCKS
    # =============================================
    # reset parent_items so search/filter works correctly
    if 'parent_items' in globals():
        parent_items.clear()
    else:
        globals()['parent_items'] = []

    global area_order
    area_order = list(REPORT_FORM_MAPPING.keys())
    for area_name in area_order:
        area_states[area_name] = {
            "visible": False,
            "buttons": [],
            "button": None
        }
        # create a block frame for each area to allow pack_forget/pack when filtering
        block_frame = tk.Frame(area_scrollable_frame, bg="white")
        list_button = tk.Button(
            block_frame,
            text=area_name,
            font=("Arial", 12, "bold"),
            height=2,
            width=18,
            bg="SystemButtonFace"
        )
        list_button.config(command=lambda a=area_name: toggle_area(a))
        list_button.pack(fill="x", padx=5, pady=4)
        # pack the block into the scrollable frame
        block_frame.pack(fill="x", padx=5, pady=4)
        area_states[area_name]["button"] = list_button
        parent_items.append(
    (block_frame, list_button)
)
    # =============================================
    # UPDATE SCROLL REGION
    # =============================================
    area_scrollable_frame.update_idletasks()
    area_canvas.configure(
        scrollregion=area_canvas.bbox("all")
    )
# ==== SAO CHÉP VĂN BẢN ====
def copy_text_to_clipboard():
    root.clipboard_clear()
    text = output_text.get("1.0", tk.END)
    root.clipboard_append(text)
    root.update()

# ==== XOÁ VĂN BẢN ====
def clear_text_output():
    output_text.config(state='normal')
    output_text.delete("1.0", tk.END)
    output_text.config(state='disabled')
    reset_timer()

# ==== TÔ MÀU TIẾN TRÌNH ====
def handle_first_box_fill():
    global first_box_filled
    if not first_box_filled:
        boxes[0].config(bg="green")
        box_colors[0] = "green"
        box_filled[0] = True
        first_box_filled = True
        update_hint("Đã ghi nhận sự cố, tiến hành báo cáo lên group chung và tiếp tục theo dõi sự cố đang diễn ra...")
        return True
    return False

def fill_box(index):
    if index == 0 or box_filled[index - 1]:
        boxes[index].config(bg="green")
        box_colors[index] = "green"
        box_filled[index] = True
        return True
    else:
        messagebox.showwarning("Chưa hoàn tất bước trước", f"Vui lòng hoàn thành bước {index} trong quy trình xử lý sự cố trước khi tiếp tục.")
        return False

def make_cmd(fid, b, fname, is_no_error=False):
    def cmd():
        set_active_child_button(b)
        handle_first_box_fill()  # ✅ Bây giờ sẽ chỉ chạy khi bấm
        if is_no_error:
            show_text_from_local(fid, fname, is_no_error=True, start_timer_flag=False)
        else:
            show_text_from_local(fid, fname, start_timer_flag=True)
    return cmd

# ==== khu vực tạo các cửa sổ chức năng ================================================================================================================================
# == Cửa sổ contact ==
# =========================================================
# CREATE CONTACT WINDOW
# =========================================================
def create_new_window_contact(title,content=None):
    # =====================================================
    # WINDOW
    # =====================================================
    new_window = tk.Toplevel(root)
    new_window.title(title)
    new_window.geometry("600x450")
    new_window.configure(bg="white")
    new_window.transient(root)
    new_window.grab_set()

    # =====================================================
    # STATE
    # =====================================================
    confirm_var = tk.StringVar(
    master=new_window,
    value="not_confirmed"
    )

    def on_confirm_change(*args):
        print("Selected:", confirm_var.get())
        toggle_entry_fields()

    confirm_var.trace_add(
        "write",
        on_confirm_change
    )


    CONTACT_SAMPLE_KEYWORD = "CONTACT_FORM"

    # =====================================================
    # CONFIRM FRAME
    # =====================================================
    confirm_frame = tk.LabelFrame(
        new_window,
        text="Tình trạng confirm",
        font=("Arial", 12, "bold"),
        bg="white"
    )
    confirm_frame.pack(
        padx=20,
        pady=10,
        fill="x"
    )
    # =====================================================
    # FORM FRAME
    # =====================================================
    form_frame = tk.Frame(
        new_window,
        bg="white"
    )
    form_frame.pack(
        padx=20,
        pady=10,
        fill="x"
    )
    # =====================================================
    # DEPARTMENT
    # =====================================================
    tk.Label(
        form_frame,
        text="Tên bộ phận:",
        font=("Arial", 11),
        bg="white"
    ).grid(
        row=0,
        column=0,
        sticky="w",
        pady=5
    )

    dept_entry = tk.Entry(
        form_frame,
        font=("Arial", 11)
    )

    dept_entry.grid(
        row=0,
        column=1,
        pady=5,
        sticky="ew"
    )
    # =====================================================
    # DEVICE
    # =====================================================
    tk.Label(
        form_frame,
        text="Tên thiết bị:",
        font=("Arial", 11),
        bg="white"
    ).grid(
        row=1,
        column=0,
        sticky="w",
        pady=5
    )

    device_entry = tk.Entry(
        form_frame,
        font=("Arial", 11)
    )

    device_entry.grid(
        row=1,
        column=1,
        pady=5,
        sticky="ew"
    )

    # =====================================================
    # STATUS
    # =====================================================
    tk.Label(
        form_frame,
        text="Tình trạng:",
        font=("Arial", 11),
        bg="white"
    ).grid(
        row=2,
        column=0,
        sticky="w",
        pady=5
    )
    status_entry = ttk.Combobox(
        form_frame,
        font=("Arial", 11),
        state="readonly",
        values=[
            "Đang xử lý",
            "Đã xử lý",
            "Chờ xử lý",
            "Chờ thông tin",
            "Không chọn"
        ]
    )
    status_entry.grid(
        row=2,
        column=1,
        pady=5,
        sticky="ew"
    )
    status_entry.current(3)

    # =====================================================
    # DESCRIPTION
    # =====================================================
    tk.Label(
        form_frame,

        text="Mô tả:",

        font=("Arial", 11),

        bg="white"
    ).grid(
        row=3,
        column=0,
        sticky="nw",
        pady=5
    )

    desc_entry = tk.Text(
        form_frame,
        font=("Arial", 11),
        height=6,
        width=40
    )

    desc_entry.grid(
        row=3,
        column=1,
        pady=5,
        sticky="ew"
    )

    form_frame.columnconfigure(
        1,
        weight=1
    )
    # =====================================================
    # TOGGLE ENTRY FIELDS
    # =====================================================

    def toggle_entry_fields():
        is_not_confirmed = (
            confirm_var.get() == "not_confirmed"
        )
        # =================================================
        # ENABLE
        # =================================================
        if is_not_confirmed:

            dept_entry.config(
                state="normal"
            )

            device_entry.config(
                state="normal"
            )

            # combobox phải readonly
            status_entry.config(
                state="readonly"
            )

            desc_entry.config(
                state="normal"
            )

        # =================================================
        # DISABLE
        # =================================================
        else:
            dept_entry.config(
                state="disabled"
            )
            device_entry.config(
                state="disabled"
            )
            status_entry.config(
                state="disabled"
            )
            desc_entry.config(
                state="disabled"
            )
    new_window.update_idletasks()
    # =====================================================
    # RADIO BUTTONS
    # =====================================================
    rb_confirm = tk.Radiobutton(
    confirm_frame,
    text="Đã confirm",variable=confirm_var,
    value="confirmed",
    bg="white",font=("Arial", 11)
    )

    rb_not_confirm = tk.Radiobutton(
        confirm_frame,
        text="Chưa confirm",variable=confirm_var,
        value="not_confirmed",
        bg="white",font=("Arial", 11)
    )

    rb_confirm.pack(anchor="w",padx=10,pady=2)
    rb_not_confirm.pack(anchor="w",padx=10,pady=2)
    # =====================================================
    # INITIAL STATE
    # =====================================================
    toggle_entry_fields()
    # =====================================================
    # HANDLE OK
    # =====================================================
    def handle_ok():
        # =============================================
        # CONFIRMED
        # =============================================
        if confirm_var.get() != "not_confirmed":
            new_window.destroy()
            return
        # =============================================
        # GET USER INPUT
        # =============================================
        dept = dept_entry.get().strip()
        device = device_entry.get().strip()
        status_val = status_entry.get().strip()
        desc = desc_entry.get(
            "1.0",
            tk.END
        ).strip()
        try:
            # =========================================
            # FIND TEMPLATE
            # =========================================
            target_file = None
            for filename in os.listdir(
                REPORT_FORM_DIR
            ):
                if (
                    CONTACT_SAMPLE_KEYWORD
                    in filename.upper()
                ):
                    target_file = os.path.join(
                        REPORT_FORM_DIR,
                        filename
                    )
                    break
            # =========================================
            # NOT FOUND
            # =========================================
            if not target_file:
                raise FileNotFoundError(
                    f"Không tìm thấy "
                    f"'{CONTACT_SAMPLE_KEYWORD}'"
                )
            # =========================================
            # READ TEMPLATE
            # =========================================
            with open(
                target_file,"r",encoding="utf-8",errors="ignore"
            ) as f:
                lines = f.readlines()
            # =========================================
            # REPLACE PLACEHOLDER
            # =========================================
            replaced_lines = []
            for line in lines:
                original_line = line
                line = line.replace("[title]",dept)
                line = line.replace("[department]",dept)
                line = line.replace("[device]",device)
                line = line.replace("[status]",status_val)
                line = line.replace("[description]",desc)
                stripped_line = line.strip()
                # =====================================
                # SKIP EMPTY FIELD
                # =====================================
                if (
                    (
                        "[title]"
                        in original_line
                        and not dept
                    )
                    or
                    (
                        "[department]"
                        in original_line
                        and not dept
                    )
                    or
                    (
                        "[device]"
                        in original_line
                        and not device
                    )
                    or
                    (
                        "[status]"
                        in original_line
                        and (
                            not status_val
                            or status_val == "Không chọn"
                        )
                    )
                    or
                    (
                        "[description]"
                        in original_line
                        and not desc
                    )
                ):
                    continue
                if not stripped_line:
                    continue
                replaced_lines.append(stripped_line)
            # =========================================
            # FINAL CONTENT
            # =========================================
            content = "\n".join(replaced_lines)
        except Exception as e:
            content = (
                f"Lỗi khi xử lý "
                f"CONTACT_FORM:\n{e}"
            )
        # =============================================
        # SHOW CONTENT
        # =============================================
        output_text.config(state="normal")
        output_text.delete("1.0",tk.END)
        output_text.insert(tk.END,content)
        output_text.config(state="disabled")
        # =============================================
        # START TIMER
        # =============================================
        try:
            if fill_box(1):
                start_timer()
        except:
            pass
        # =============================================
        # CLOSE WINDOW
        # =============================================
        new_window.destroy()
    # =====================================================
    # OK BUTTON
    # =====================================================
    ok_button = tk.Button(
        new_window,
        text="OK",
        font=("Arial", 12, "bold"),
        bg="green",
        fg="white",
        width=15,
        command=handle_ok
    )
    ok_button.pack(
        pady=15
    )
# Nhớ thêm dòng này lên đầu file code của bạn nhé:
# from tkcalendar import DateEntry

# == CỬA SỔ STATUS ==
def create_new_window_status(title, content=None):
    new_window = tk.Toplevel(root)
    new_window.title(title)
    new_window.geometry("600x600")
    new_window.configure(bg="white")

    # =====================================================
    # CONFIRM STATE
    # =====================================================
    confirm_var = tk.StringVar(master=new_window, value="not_confirmed")

    def on_confirm_change(*args):
        print("Selected:", confirm_var.get())
        toggle_entry_fields()

    confirm_var.trace_add("write", on_confirm_change)

    # =====================================================
    # CONFIRM FRAME
    # =====================================================
    confirm_frame = tk.LabelFrame(new_window, text="Đã confirm chưa?", font=("Arial", 12, "bold"))
    confirm_frame.pack(padx=20, pady=10, fill="x")

    # =====================================================
    # FORM FRAME
    # =====================================================
    form_frame = tk.Frame(new_window, bg="white")
    form_frame.pack(padx=20, pady=10, fill="x")

    # =====================================================
    # DEPARTMENT
    # =====================================================
    tk.Label(form_frame, text="Tên bộ phận:", font=("Arial", 11), bg="white").grid(row=0, column=0, sticky="w", pady=5)
    dept_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    dept_entry.grid(row=0, column=1, pady=5, sticky="ew")

    # =====================================================
    # DEVICE
    # =====================================================
    tk.Label(form_frame, text="Tên thiết bị:", font=("Arial", 11), bg="white").grid(row=1, column=0, sticky="w", pady=5)
    device_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    device_entry.grid(row=1, column=1, pady=5, sticky="ew")

    # =====================================================
    # STATUS
    # =====================================================
    tk.Label(form_frame, text="Tình trạng:", font=("Arial", 11), bg="white").grid(row=2, column=0, sticky="w", pady=5)
    status_entry = ttk.Combobox(
        form_frame,
        font=("Arial", 11),
        state="disabled",
        values=[
            "Alarm - Chưa xử lý",
            "Alarm - Đã xử lý",
            "Alarm - Chờ xử lý",
            "Normal - Đã xử lý",
            "Normal - Chờ xử lý",
            "Không chọn"
        ]
    )
    status_entry.grid(row=2, column=1, pady=5, sticky="ew")
    status_entry.set("Không chọn")

    # =====================================================
    # START DATE & TIME (SỬ DỤNG DATEENTRY)
    # =====================================================
    tk.Label(form_frame, text="Ngày bắt đầu:", font=("Arial", 11), bg="white").grid(row=3, column=0, sticky="w", pady=5)
    start_date_entry = DateEntry(
        form_frame, 
        font=("Arial", 11), 
        width=12, 
        background='darkblue',
        foreground='white', 
        borderwidth=2, 
        date_pattern='dd/mm/yyyy', 
        state="disabled"
    )
    start_date_entry.grid(row=3, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Thời gian bắt đầu (HH:MM):", font=("Arial", 11), bg="white").grid(row=4, column=0, sticky="w", pady=5)
    start_time_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    start_time_entry.grid(row=4, column=1, pady=5, sticky="ew")

    # =====================================================
    # END DATE & TIME (SỬ DỤNG DATEENTRY)
    # =====================================================
    tk.Label(form_frame, text="Ngày kết thúc:", font=("Arial", 11), bg="white").grid(row=5, column=0, sticky="w", pady=5)
    end_date_entry = DateEntry(
        form_frame, 
        font=("Arial", 11), 
        width=12, 
        background='darkblue',
        foreground='white', 
        borderwidth=2, 
        date_pattern='dd/mm/yyyy', 
        state="disabled"
    )
    end_date_entry.grid(row=5, column=1, pady=5, sticky="ew")

    tk.Label(form_frame, text="Thời gian kết thúc (HH:MM):", font=("Arial", 11), bg="white").grid(row=6, column=0, sticky="w", pady=5)
    end_time_entry = tk.Entry(form_frame, font=("Arial", 11), state="disabled")
    end_time_entry.grid(row=6, column=1, pady=5, sticky="ew")

    # =====================================================
    # DESCRIPTION
    # =====================================================
    tk.Label(form_frame, text="Mô tả:", font=("Arial", 11), bg="white").grid(row=7, column=0, sticky="nw", pady=5)
    desc_entry = tk.Text(form_frame, font=("Arial", 11), height=5, width=40, state="disabled")
    desc_entry.grid(row=7, column=1, pady=5, sticky="ew")

    form_frame.columnconfigure(1, weight=1)

    # =====================================================
    # ENABLE / DISABLE FIELD
    # =====================================================
    def toggle_entry_fields():
        is_not_confirmed = (confirm_var.get() == "not_confirmed")
        
        state_normal = "normal" if is_not_confirmed else "disabled"
        state_readonly = "readonly" if is_not_confirmed else "disabled"

        dept_entry.config(state=state_normal)
        device_entry.config(state=state_normal)
        # tkcalendar dùng readonly để người dùng chỉ bấm chọn lịch chứ không gõ tay bậy bạ vào được
        start_date_entry.config(state="readonly" if is_not_confirmed else "disabled") 
        start_time_entry.config(state=state_normal)
        end_date_entry.config(state="readonly" if is_not_confirmed else "disabled")
        end_time_entry.config(state=state_normal)
        status_entry.config(state=state_readonly)
        desc_entry.config(state=state_normal)

    new_window.update_idletasks()

    # =====================================================
    # RADIO BUTTON
    # =====================================================
    rb_confirm = tk.Radiobutton(confirm_frame, text="Đã confirm", variable=confirm_var, value="confirmed", bg="white", font=("Arial", 11))
    rb_not_confirm = tk.Radiobutton(confirm_frame, text="Chưa confirm", variable=confirm_var, value="not_confirmed", bg="white", font=("Arial", 11))
    
    rb_confirm.pack(anchor="w", padx=10, pady=2)
    rb_not_confirm.pack(anchor="w", padx=10, pady=2)

    toggle_entry_fields()

    # =====================================================
    # TEMPLATE KEYWORD
    # =====================================================
    CONFIRM_SAMPLE_KEYWORD = "CONFIRM_FORM"

    # =====================================================
    # HANDLE OK
    # =====================================================
    def handle_ok():
        if confirm_var.get() != "not_confirmed":
            new_window.destroy()
            return

        # =============================================
        # GET INPUT
        # =============================================
        dept = dept_entry.get().strip()
        device = device_entry.get().strip()
        status_val = status_entry.get().strip()
        if status_val == "Không chọn":
            status_val = ""

        # Lấy giá trị chuỗi ngày theo định dạng dd/mm/yyyy từ bảng lịch
        start_date_str = start_date_entry.get() 
        start_time_str = start_time_entry.get().strip()
        end_date_str = end_date_entry.get()
        end_time_str = end_time_entry.get().strip()
        desc = desc_entry.get("1.0", tk.END).strip()

        # =============================================
        # CALCULATE PROCESS TIME (Only if time is provided)
        # =============================================
        time_process = ""
        if start_time_str and end_time_str:
            try:
                fmt = "%d/%m/%Y %H:%M"
                start_dt = datetime.datetime.strptime(f"{start_date_str} {start_time_str}", fmt)
                end_dt = datetime.datetime.strptime(f"{end_date_str} {end_time_str}", fmt)

                diff_minutes = int((end_dt - start_dt).total_seconds() / 60)

                # Cảnh báo nếu chọn ngày kết thúc bé hơn ngày bắt đầu
                if diff_minutes < 0:
                    messagebox.showwarning("Lỗi logic", "Thời gian kết thúc không thể nhỏ hơn thời gian bắt đầu!")
                    return

                time_process = f"{diff_minutes} phút ({start_time_str} {start_date_str} - {end_time_str} {end_date_str})"
            
            except ValueError:
                messagebox.showwarning("Sai định dạng giờ", "Vui lòng nhập đúng định dạng giờ HH:MM (VD: 09:30)")
                return

        # =========================================
        # FIND & READ TEMPLATE
        # =========================================
        try:
            target_file = None
            for filename in os.listdir(REPORT_FORM_DIR):
                if CONFIRM_SAMPLE_KEYWORD in filename.upper():
                    target_file = os.path.join(REPORT_FORM_DIR, filename)
                    break

            if not target_file:
                raise FileNotFoundError(f"Không tìm thấy file '{CONFIRM_SAMPLE_KEYWORD}'")

            with open(target_file, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()

            # =========================================
            # REPLACE PLACEHOLDER 
            # Dùng DELETE_THIS_LINE nếu trường bị trống
            # =========================================
            marker = "DELETE_THIS_LINE"
            replacements = {
                "[tilte]": dept if dept else marker,
                "[title]": dept if dept else marker,
                "[department]": dept if dept else marker,
                "[device]": device if device else marker,
                "[status]": status_val if status_val else marker,
                "[time_process]": time_process if time_process else marker,
                "[description]": desc if desc else marker
            }

            for key, value in replacements.items():
                content = content.replace(key, value)

            # =========================================
            # XÓA CÁC DÒNG CHỨA DỮ LIỆU TRỐNG
            # =========================================
            cleaned_lines = []
            for line in content.splitlines():
                if marker in line:
                    continue  
                if line.strip():
                    cleaned_lines.append(line.strip())

            content = "\n".join(cleaned_lines)

        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi xử lý file:\n{e}")
            return

        # =============================================
        # SHOW CONTENT
        # =============================================
        output_text.config(state="normal")
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, content)
        output_text.config(state="disabled")

        # =============================================
        # FILL BOX & CLOSE WINDOW
        # =============================================
        fill_box(2)
        new_window.destroy()

    # =====================================================
    # OK BUTTON
    # =====================================================
    ok_button = tk.Button(new_window, text="OK", font=("Arial", 12, "bold"), bg="green", fg="white", command=handle_ok)
    ok_button.pack(pady=10)

# == Cửa sổ note ==
schedule_running = False
schedule_after_id = None
def create_new_window_note():
    # Thư mục lưu dữ liệu
    DATA_DIR = NOTE_ARCHIVE_DIR
    os.makedirs(DATA_DIR, exist_ok=True)
    # === Luồng kế hoạch ===
    def run_schedule():
        global schedule_after_id
        try:
            if not root.winfo_exists():
                return
            schedule.run_pending()
            schedule_after_id = root.after(
                1000,
                run_schedule
            )
        except tk.TclError:
            schedule_after_id = None

    global schedule_running
    if not schedule_running:
        schedule_running = True
        run_schedule()

    # === Tạo Note ===
    def get_next_stt():
        used_numbers = []
        for filename in os.listdir(DATA_DIR):
            if filename.startswith("reminders") and filename.endswith(".json"):
                try:
                    number = int(filename.replace("reminders", "").replace(".json", ""))
                    used_numbers.append(number)
                except:
                    continue
        count = 1
        while count in used_numbers:
            count += 1
        return count

    def update_stt_label():
        current_stt.set(str(len([f for f in os.listdir(DATA_DIR) if f.endswith(".json")])))

    def save_reminder_to_new_file(reminder_data):
        stt = get_next_stt()
        file_path = os.path.join(DATA_DIR, f"reminders{stt}.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(reminder_data, f, ensure_ascii=False, indent=4)
        update_stt_label()

    def schedule_reminder(keyword, content, times, days, months, mode, file_path=None, delete_mode="delete"):
        for t in times:
            print("Reminder fired:",keyword,datetime.datetime.now())
            def job(t=t):
                now = datetime.datetime.now()
                if str(now.day) in days and str(now.month) in months:
                    # Hiển thị thông báo đúng luồng giao diện
                    def show_popup():
                        messagebox.showinfo(f"Thông báo: {keyword}", f"[{t}] {content}")

                        if mode == "1 lần" and file_path and os.path.exists(file_path):
                            try:
                                with open(file_path, "r", encoding="utf-8") as f:
                                    data = json.load(f)

                                # ✅ Nếu chọn delete → xóa file
                                if delete_mode == "delete":
                                    os.remove(file_path)
                                else:
                                    # ✅ Nếu chọn keep → chỉ update "done": True
                                    if isinstance(data, dict):
                                        data["done"] = True
                                        with open(file_path, "w", encoding="utf-8") as f:
                                            json.dump(data, f, ensure_ascii=False, indent=4)
                            except Exception as e:
                                print(f"Lỗi xử lý file {file_path}: {e}")
                    root.after(
                        0,
                        show_popup
                    )

                    if mode == "1 lần":
                        return schedule.CancelJob

            schedule.every().day.at(t).do(job)

    def add_reminder():
        keyword = keyword_entry.get().strip()
        content = content_entry.get().strip()
        time_input = time_entry.get().strip()
        day_input = day_entry.get().strip()
        month_input = month_entry.get().strip()
        mode = intensity_var.get()

        # ==== Chuẩn hóa thời gian ====
        time_strs = time_input.split(",")
        normalized_times = []
        for t in time_strs:
            t = t.strip()
            if not t:
                continue
            try:
                h, m = map(int, t.split(":"))
                normalized_time = f"{h:02d}:{m:02d}"
                # Kiểm tra hợp lệ bằng datetime
                datetime.datetime.strptime(normalized_time, "%H:%M")
                normalized_times.append(normalized_time)
            except:
                messagebox.showerror("Lỗi", f"Thời gian không hợp lệ: {t}", parent=note_window)
                return

        # ==== Ngày ====
        if day_input.strip().lower() == "all":
            day_strs = [str(d) for d in range(1, 32)]
        else:
            day_strs = []
            for d in day_input.split(","):
                d = d.strip()
                if not d:
                    continue
                try:
                    val = int(d)
                    assert 1 <= val <= 31
                    day_strs.append(str(val))
                except:
                    messagebox.showerror("Lỗi", f"Ngày không hợp lệ: {d}", parent=note_window)
                    return

        # ==== Tháng ====
        if month_input.strip().lower() == "all":
            month_strs = [str(m) for m in range(1, 13)]
        else:
            month_strs = []
            for m in month_input.split(","):
                m = m.strip()
                if not m:
                    continue
                try:
                    val = int(m)
                    assert 1 <= val <= 12
                    month_strs.append(str(val))
                except:
                    messagebox.showerror("Lỗi", f"Tháng không hợp lệ: {m}", parent=note_window)
                    return
        mode = intensity_var.get()

        try:
            for t in time_strs:
                datetime.datetime.strptime(t.strip(), "%H:%M")
            for d in day_strs:
                d = int(d.strip())
                assert 1 <= d <= 31
            for m in month_strs:
                m = int(m.strip())
                assert 1 <= m <= 12
        except:
            messagebox.showerror("Lỗi", "Thời gian, ngày hoặc tháng không hợp lệ", parent=note_window)
            return

        times = normalized_times
        days = [d.strip() for d in day_strs]
        months = [m.strip() for m in month_strs]

        reminder_data = {
            "keyword": keyword,                     #  Từ khóa nhắc
            "content": content,                     #  Nội dung của nhắc
            "times": times,                         #  Dữ liệu giờ phút (HH:MM)
            "days": days,                           #  Dữ liệu ngày
            "months": months,                       #  Dữ liệu tháng 
            "mode": mode,                           #  Loai nhắc ("1 lần" hoặc "Cố định")
            "delete_mode": delete_mode_var.get(),   #  thêm lựa chọn delete/keep (Chỉ dành cho note nhắc 1 lần)
            "done": False                           #  đánh dấu đã nhắc hay chưa
        }
        save_reminder_to_new_file(reminder_data)
        file_path = os.path.join(DATA_DIR, f"reminders{get_next_stt()-1}.json")
        schedule_reminder(keyword, content, times, days, months, mode, file_path, delete_mode_var.get())
        messagebox.showinfo("Thành công", f"Đã tạo note {get_next_stt()-1}.json", parent=note_window)

    def set_placeholder(entry, text):
        entry.insert(0, text)
        entry.config(fg="gray")

        def on_focus_in(event):
            if entry.get() == text:
                entry.delete(0, tk.END)
                entry.config(fg="black")
        def on_focus_out(event):
            if not entry.get():
                entry.insert(0, text)
                entry.config(fg="gray")

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def load_all_json_files():
        if not os.path.exists(DATA_DIR):
            messagebox.showerror("Lỗi", f"Không tìm thấy thư mục: {DATA_DIR}", parent=note_window)
            return []

        all_data = []
        for filename in os.listdir(DATA_DIR):
            if filename.endswith(".json"):
                file_path = os.path.join(DATA_DIR, filename)
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if isinstance(data, dict) and "keyword" in data:
                            data["_file"] = file_path
                            all_data.append(data)
                        elif isinstance(data, list):
                            for item in data:
                                if isinstance(item, dict) and "keyword" in item:
                                    item["_file"] = file_path
                                    if "delete_mode" not in item:
                                        item["delete_mode"] = "delete"
                                    if "done" not in item:
                                        item["done"] = False
                                    all_data.append(item)
                except Exception as e:
                    print(f"Lỗi đọc {filename}: {e}")
        return all_data

    def display_data(data_list):
        for row in tree.get_children():
            tree.delete(row)

        # Sort lại theo tên file reminders{n}.json để đảm bảo đúng STT thực tế
        data_list_sorted = sorted(data_list, key=lambda d: int(os.path.splitext(os.path.basename(d.get("_file", "reminders0.json")))[0].replace("reminders", "")))

        now = datetime.datetime.now()
        for i, item in enumerate(data_list_sorted, start=1):
            mode = item.get("mode", "")
            times = item.get("times", [])
            days = item.get("days", [])
            months = item.get("months", [])

            tag = ""
            if mode == "1 lần":
                expired = True
                for m in months:
                    for d in days:
                        for t in times:
                            try:
                                h, mn = map(int, t.split(":"))
                                scheduled_time = datetime.datetime(year=now.year, month=int(m), day=int(d), hour=h, minute=mn)
                                if scheduled_time >= now:
                                    expired = False
                                    break
                            except:
                                pass
                tag = "one_time_valid" if not expired else "one_time_expired"
            else:
                tag = "recurring"

            tree.insert("", tk.END, values=(
                i,
                item["keyword"],
                item["content"],
                ", ".join(item["times"]),
                ", ".join(item["days"]),
                ", ".join(item["months"]),
                mode
            ), tags=(tag,))

    def search_data():
        keyword = search_var.get().lower()
        filtered = [item for item in full_data if keyword in item["keyword"].lower() or keyword in item["content"].lower()]
        display_data(filtered)

    def refresh_data():
        global full_data
        full_data = load_all_json_files()
        display_data(full_data)
        update_stt_label()  # cập nhật STT ghi chú tiếp theo

    def delete_selected_notes():
        global full_data
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn ít nhất một ghi chú để xóa.", parent=note_window)
            return

        confirm = messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa các ghi chú đã chọn?", parent=note_window)
        if not confirm:
            return

        to_delete = []
        deleted_files = set()

        for item_id in selected_items:
            values = tree.item(item_id, "values")
            keyword = values[1]
            content = values[2]
            for data_item in full_data:
                if data_item["keyword"] == keyword and data_item["content"] == content:
                    file_path = data_item.get("_file")
                    if file_path and os.path.exists(file_path):
                        deleted_files.add(file_path)
                    to_delete.append(data_item)
                    break

        for file_path in deleted_files:
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"Lỗi khi xóa file {file_path}: {e}")

        full_data = [item for item in full_data if item not in to_delete]
        display_data(full_data)
        messagebox.showinfo("Thành công", f"Đã xóa {len(to_delete)} ghi chú.", parent=note_window)

    if schedule_after_id:
        root.after_cancel(schedule_after_id)

    # === Giao diện chính ===
    note_window = tk.Toplevel()
    run_schedule()

    note_window.title("Trình quản lý ghi chú định kỳ")
    note_window.geometry("1000x500")

    btn_frame = tk.Frame(note_window)
    btn_frame.pack(pady=10)

    main_frame = tk.Frame(note_window)
    main_frame.pack(fill="both", expand=True)

    def show_create_note():
        for w in main_frame.winfo_children():
            w.destroy()

        update_stt_label()
        tk.Label(main_frame, text="Số ghi chú hiện tại:").pack()
        tk.Label(main_frame, textvariable=current_stt, font=("Arial", 14, "bold"), fg="blue").pack(pady=(0, 10))

        global keyword_entry, content_entry, time_entry, day_entry, month_entry, intensity_var, delete_mode_var, delete_frame

        tk.Label(main_frame, text="Từ khóa:").pack()
        keyword_entry = tk.Entry(main_frame)
        keyword_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Nội dung:").pack()
        content_entry = tk.Entry(main_frame)
        content_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Thời gian báo (HH:MM, cách nhau dấu phẩy):").pack()
        time_entry = tk.Entry(main_frame)
        set_placeholder(time_entry, "08:00,12:00,14:00,...")
        time_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Ngày báo (VD: 1,15,28):").pack()
        day_entry = tk.Entry(main_frame)
        set_placeholder(day_entry, "1,15 hoặc All")
        day_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Tháng báo (VD: 1,6,12):").pack()
        month_entry = tk.Entry(main_frame)
        set_placeholder(month_entry, "1,6,12 hoặc All")
        month_entry.pack(fill="x", padx=10)

        tk.Label(main_frame, text="Cường độ báo:").pack()
        intensity_var = tk.StringVar(value="1 lần")
        mode_combo = ttk.Combobox(main_frame, textvariable=intensity_var, values=["1 lần", "Cố định"])
        mode_combo.pack(fill="x", padx=10)

        # --- Frame chứa tick chọn ---
        delete_frame = tk.LabelFrame(main_frame, text="Tùy chọn khi đã nhắc (chỉ cho loại nhắc 1 lần", padx=5, pady=5)
        delete_frame.pack(fill="x", padx=10, pady=5)

        delete_mode_var = tk.StringVar(value="delete")  # mặc định là xóa sau nhắc

        rb_delete = tk.Radiobutton(
            delete_frame, text=" Xóa khi đã nhắc", variable=delete_mode_var, value="delete"
        )
        rb_keep = tk.Radiobutton(
            delete_frame, text=" Không xóa khi đã nhắc", variable=delete_mode_var, value="keep"
        )

        rb_delete.pack(side="left", padx=5)
        rb_keep.pack(side="left", padx=5)

        # Hàm enable/disable frame dựa trên mode
        def update_delete_frame_state(*args):
            if intensity_var.get() == "1 lần":
                for child in delete_frame.winfo_children():
                    child.configure(state="normal")
            else:
                for child in delete_frame.winfo_children():
                    child.configure(state="disabled")

        # Gán sự kiện thay đổi mode
        intensity_var.trace_add("write", update_delete_frame_state)

        # gọi 1 lần ban đầu để set trạng thái đúng
        update_delete_frame_state()

        tk.Button(main_frame, text="Thêm Nhắc", command=add_reminder).pack(pady=15)

    def show_view_notes():
        for w in main_frame.winfo_children():
            w.destroy()

        search_frame = tk.Frame(main_frame)
        search_frame.pack(padx=10, pady=(10, 0), fill=tk.X)

        tk.Label(search_frame, text="Tìm kiếm:").pack(side=tk.LEFT, padx=(0, 5))

        global search_var, tree, full_data
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(search_frame, text="Tìm", command=search_data,
                  bg="#00ccff", fg="white", activebackground="#006699", activeforeground="white").pack(side=tk.LEFT, padx=5)

        tk.Button(search_frame, text="Làm mới", command=refresh_data,
                  bg="#00cc66", fg="white", activebackground="#006600", activeforeground="white").pack(side=tk.LEFT, padx=5)

        tk.Button(search_frame, text="Xóa ghi chú đã chọn", command=delete_selected_notes,
                  bg="#cc3300", fg="white", activebackground="#990000", activeforeground="white").pack(side=tk.LEFT, padx=5)

        search_entry.bind("<Return>", lambda event: search_data())

        table_frame = tk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        columns = ("STT", "Từ khóa", "Nội dung", "Thời gian", "Ngày", "Tháng", "Cường độ")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", yscrollcommand=scrollbar.set, selectmode="extended")
        tree.tag_configure("one_time_valid", background="#d4fcd4")     # xanh lá
        tree.tag_configure("one_time_expired", background="#f8d4d4")   # đỏ
        tree.tag_configure("recurring", background="#d4eaff")          # xanh da trời
        scrollbar.config(command=tree.yview)
        # Cài đặt tiêu đề + cột
        tree.heading("STT", text="STT")
        tree.column("STT", width=10, anchor="center")

        tree.heading("Từ khóa", text="Từ khóa")
        tree.column("Từ khóa", width=60, anchor="center")

        tree.heading("Nội dung", text="Nội dung")
        tree.column("Nội dung", width=400, anchor="w")

        tree.heading("Thời gian", text="Thời gian")
        tree.column("Thời gian", width=50, anchor="center")

        tree.heading("Ngày", text="Ngày")
        tree.column("Ngày", width=100, anchor="center")

        tree.heading("Tháng", text="Tháng")
        tree.column("Tháng", width=100, anchor="center")

        tree.heading("Cường độ", text="Cường độ")
        tree.column("Cường độ", width=100, anchor="center")

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center")

        tree.column("Nội dung", width=200, anchor="w")
        tree.pack(fill=tk.BOTH, expand=True)

        full_data = load_all_json_files()
        display_data(full_data)

    tk.Button(btn_frame, text="Tạo Note", width=20, command=show_create_note).pack(side="left", padx=10)
    tk.Button(btn_frame, text="Xem Note", width=20, command=show_view_notes).pack(side="left", padx=10)

    current_stt = tk.StringVar()
    show_create_note()

    # ==== Lên lịch lại tất cả ghi chú đã lưu ====
    for reminder in load_all_json_files():
        if reminder.get("done", False):
            continue
        schedule_reminder(
            reminder["keyword"],
            reminder["content"],
            reminder["times"],
            reminder["days"],
            reminder["months"],
            reminder["mode"],
            reminder.get("_file"),
            reminder.get("delete_mode", "delete")
        )
# == Cửa sổ hình ảnh Daviteq ==
def create_new_window_image_daviteq(title):
    # =====================================================
    # BUILD IMAGE MAPPING FROM LOCAL
    # =====================================================
    # =====================================================
    # BUILD IMAGE MAPPING FROM LOCAL (ĐỒNG BỘ HAI NHÁNH THƯ MỤC)
    # =====================================================
    def build_image_mapping():
        # Cấu hình danh sách thư mục cần quét cho từng danh mục cha
        image_roots = {
            "LAYOUT": [AEONGMS_IMAGE_LAYOUT_ARCHIVE_DIR, MAXVALU_IMAGE_LAYOUT_ARCHIVE_DIR],
            "GATEWAY": [AEONGMS_IMAGE_GATEWAY_ARCHIVE_DIR, MAXVALU_IMAGE_GATEWAY_ARCHIVE_DIR],
            "SENSOR": [AEONGMS_IMAGE_SENSOR_ARCHIVE_DIR, MAXVALU_IMAGE_SENSOR_ARCHIVE_DIR],
            "ALARMPOINT": [AEONGMS_IMAGE_AL_ARCHIVE_DIR, MAXVALU_IMAGE_AL_ARCHIVE_DIR]
        }

        result = {}
        for category, folders_list in image_roots.items():
            result[category] = {}
            
            # Duyệt qua từng thư mục cấu hình (quét cả GMS lẫn Maxvalu)
            for folder in folders_list:
                if not os.path.exists(folder):
                    continue
                for filename in os.listdir(folder):
                    filepath = os.path.join(folder, filename)
                    if not os.path.isfile(filepath):
                        continue
                        
                    # Bỏ extension (.png, .jpg...)
                    name_without_ext = os.path.splitext(filename)[0]
                    # Split theo ký tự "_"
                    parts = name_without_ext.split("_")
                    
                    # Định dạng chuẩn: NVL_DELICA hoặc TQB_BAKERY
                    if len(parts) < 2:
                        continue
                        
                    area = parts[0].upper()
                    device = "_".join(parts[1:])
                    
                    if area not in result[category]:
                        result[category][area] = []
                        
                    result[category][area].append({
                        "device": device,
                        "path": filepath,
                        "filename": filename
                    })
                    
        # Sắp xếp lại thứ tự thiết bị (Device) theo Alphabet để hiển thị đẹp mắt
        for category in result:
            for area in result[category]:
                result[category][area].sort(
                    key=lambda x: x["device"]
                )
        return result
    # =====================================================
    # SHOW IMAGES
    # =====================================================
    def show_images(image_list):
        for widget in image_frame.winfo_children():
            widget.destroy()
        for idx, item in enumerate(image_list):
            try:
                img_path = item["path"]
                device_name = item["device"]
                img = Image.open(img_path)
                img.thumbnail(
                    (140, 100),
                    Image.Resampling.LANCZOS
                )
                photo = ImageTk.PhotoImage(img)
                row = idx // max_columns
                col = idx % max_columns
                # =============================================
                # ITEM FRAME
                # =============================================
                item_frame = tk.Frame(
                    image_frame,
                    bg="white"
                )
                item_frame.grid(
                    row=row,
                    column=col,
                    padx=10,
                    pady=10
                )
                # =============================================
                # IMAGE BUTTON
                # =============================================
                label_img = tk.Label(
                    item_frame,
                    image=photo,
                    bg="white",
                    cursor="hand2"
                )

                label_img.image = photo
                label_img.pack()
                # =============================================
                # DEVICE LABEL
                # =============================================
                label_text = tk.Label(
                    item_frame,
                    text=device_name,
                    bg="white",
                    font=("Arial", 10, "bold"),
                    wraplength=140,
                    justify="center"
                )

                label_text.pack(
                    pady=(5, 0)
                )
                # =============================================
                # OPEN LARGE IMAGE
                # =============================================
                label_img.bind(
                    "<Button-1>",
                    lambda e, p=img_path:
                        open_large_image(p)
                )
            except Exception as e:
                print(
                    f"❌ Lỗi xử lý ảnh: {e}"
                )
    # =====================================================
    # OPEN LARGE IMAGE
    # =====================================================
    def open_large_image(img_path):
        try:
            img = Image.open(img_path)
            # scale 50%
            scale_factor = 0.5
            new_size = (
                int(img.width * scale_factor),
                int(img.height * scale_factor)
            )
            img_resized = img.resize(
                new_size,
                Image.Resampling.LANCZOS
            )
            photo = ImageTk.PhotoImage(img_resized)
            popup = tk.Toplevel(new_window)
            popup.title("DAVITEQ IMAGE DATA")
            popup.configure(bg="white")
            lbl = tk.Label(
                popup,
                image=photo,
                bg="white"
            )

            lbl.image = photo

            lbl.pack(
                padx=10,
                pady=10
            )

            btn = tk.Button(
                popup,
                text="Open Image",
                font=("Arial", 10, "bold"),
                command=lambda:
                    open_image_external(img_path)
            )

            btn.pack(
                pady=10
            )

        except Exception as e:

            print(
                f"❌ Lỗi mở ảnh lớn: {e}"
            )

    # =====================================================
    # OPEN IMAGE EXTERNAL
    # =====================================================
    def open_image_external(path):

        try:

            os.startfile(path)

        except Exception as e:

            print(
                f"❌ Lỗi mở ảnh: {e}"
            )

    # =====================================================
    # SUB BUTTON CLICK
    # =====================================================
    def on_sub_button_click(
        btn_clicked,
        image_list
    ):

        nonlocal selected_sub_button

        # reset all button
        for frame in category_frames.values():

            for widget in frame.winfo_children():

                if isinstance(widget, tk.Button):

                    widget.config(
                        bg="white",
                        fg="black"
                    )

        selected_sub_button = btn_clicked

        selected_sub_button.config(
            bg="#4CAF50",
            fg="white"
        )

        show_images(image_list)

    # =====================================================
    # TOGGLE CATEGORY
    # =====================================================
    def toggle_sub_buttons(category_name):

        nonlocal selected_parent_button

        # reset all parent button
        for btn in parent_buttons.values():

            btn.configure(
                bg="white",
                fg="black"
            )

        selected_parent_button = parent_buttons[category_name]

        selected_parent_button.configure(
            bg="#247985",
            fg="white"
        )

        # hide all frame
        for frame in category_frames.values():

            frame.pack_forget()

        # show selected frame
        category_frames[category_name].pack(
            fill="y"
        )

    # =====================================================
    # WINDOW
    # =====================================================
    new_window = tk.Toplevel()

    new_window.title(title)

    new_window.geometry("1200x650")

    new_window.configure(bg="white")

    # =====================================================
    # LEFT FRAME
    # =====================================================
    left_frame = tk.Frame(
        new_window,
        width=170,
        bg="#f0f0f0"
    )

    left_frame.pack(
        side="left",
        fill="y"
    )

    # =====================================================
    # SUB BUTTON FRAME
    # =====================================================
    sub_button_frame = tk.Frame(
        new_window,
        width=180,
        bg="#e8e8e8"
    )

    sub_button_frame.pack(side="left",fill="y"
    )
    # =====================================================
    # RIGHT FRAME
    # =====================================================
    right_frame = tk.Frame(new_window,bg="white")
    right_frame.pack(side="right",fill="both",expand=True)
    # =====================================================
    # SCROLLABLE CANVAS
    # =====================================================
    canvas = tk.Canvas(right_frame,bg="white")
    scrollbar = tk.Scrollbar(right_frame,orient="vertical",command=canvas.yview)
    scrollable_frame = tk.Frame(canvas,bg="white")
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )
    canvas.create_window(
        (0, 0),
        window=scrollable_frame,
        anchor="nw"
    )
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left",fill="both",expand=True)
    scrollbar.pack(side="right",fill="y")
    image_frame = scrollable_frame
    # =====================================================
    # VARIABLES
    # =====================================================
    category_frames = {}
    parent_buttons = {}
    selected_sub_button = None
    selected_parent_button = None
    max_columns = 4
    # =====================================================
    # LOAD LOCAL IMAGE DATA
    # =====================================================
    category_images = build_image_mapping()
    print("✅ IMAGE MAPPING:")
    print(json.dumps(
        category_images,
        indent=4,
        ensure_ascii=False
    ))
    # =====================================================
    # BUILD GUI DYNAMIC
    # =====================================================
    for category_name, areas in category_images.items():
        # =============================================
        # PARENT BUTTON
        # =============================================
        parent_btn = tk.Button(
            left_frame,
            text=category_name,
            width=15,pady=5,
            bg="white",fg="black",
            font=("Arial", 10, "bold"),
            activebackground="#e0e0e0",
            command=lambda c=category_name:
                toggle_sub_buttons(c)
        )
        parent_btn.pack(pady=(10, 0))
        parent_buttons[category_name] = parent_btn
        # =============================================
        # SUB FRAME
        # =============================================
        sub_frame = tk.Frame(sub_button_frame,bg="#e8e8e8")
        category_frames[category_name] = sub_frame
        # =============================================
        # AREA BUTTONS
        # =============================================
        for area_name, image_list in areas.items():
            sub_btn = tk.Button(
                sub_frame,
                text=area_name,
                width=15,pady=5,
                relief="raised",
                bg="white",fg="black",
                font=("Arial", 10, "bold"),
                bd=1,activebackground="#e0e0e0"
            )

            sub_btn.pack(padx=10,pady=3)
            sub_btn.config(
                command=lambda b=sub_btn,
                               il=image_list:
                    on_sub_button_click(
                        b,
                        il
                    )
            )

    # =====================================================
    # AUTO SELECT FIRST CATEGORY
    # =====================================================
    if category_images:

        first_category = list(
            category_images.keys()
        )[0]

        toggle_sub_buttons(
            first_category
        )

        # auto select first area
        first_area = list(
            category_images[first_category].keys()
        )[0]

        first_image_list = category_images[
            first_category
        ][
            first_area
        ]

        first_sub_frame = category_frames[
            first_category
        ]

        first_sub_btn = first_sub_frame.winfo_children()[0]

        on_sub_button_click(
            first_sub_btn,
            first_image_list
        )
# == Cửa sổ tài liệu ==
def create_documentary_viewer(token, share_url):
    files = list_files_from_url(token, share_url)  # Lấy file từ OneDrive Azure
    filtered_files = files.copy()

    # ==== Hàm tách tag từ tên file ====
    def extract_tags(filename):
        tags = re.findall(r'\(([^)]+)\)', filename)
        return ", ".join(tags) if tags else "Khác"

    def update_table(event=None):
        # 1. Lấy trực tiếp từ Entry thay vì StringVar để tránh độ trễ
        keyword = entry_search.get().strip().lower()

        text_widget.config(state="normal")
        text_widget.delete("1.0", tk.END)

        # 2. Thuật toán phân loại Tuple (Chắc chắn 100% đẩy lên đầu)
        def get_sort_key(file_obj):
            name = file_obj["name"].lower()
            if not keyword:
                return (0, name) # Trạng thái gốc: Xếp theo tên
            
            # Nếu tên chứa toàn bộ cụm từ khóa -> Đẩy lên Top 1 (hệ số -2)
            if keyword in name:
                return (-2, name.find(keyword), name)
                
            # Nếu tên chứa một phần từ khóa -> Đẩy lên Top 2 (hệ số -1)
            words = keyword.split()
            match_count = sum(1 for w in words if w in name)
            if match_count > 0:
                return (-1, -match_count, name)
                
            # Không liên quan -> Bỏ xuống cuối (hệ số 0)
            return (0, 0, name)

        # Sắp xếp danh sách với thuật toán mới
        sorted_files = sorted(files, key=get_sort_key)
        filtered_files.clear()
        filtered_files.extend(sorted_files)

        # Chèn Header (Cố định ở dòng 1)
        header = f"{'STT':<5}{'LOẠI':<15}{'TÊN FILE'}\n"
        text_widget.insert(tk.END, header, "header")

        # Chèn danh sách File
        for idx, file in enumerate(sorted_files, start=1):
            filename = file["name"]
            tag_label = extract_tags(filename)
            
            filepath = os.path.join(DOCUMENTARY_ARCHIVE_DIR, filename)
            downloaded = os.path.exists(filepath)
            row_tag = "downloaded" if downloaded else "not_downloaded"

            line = f"{idx:<5}{tag_label:<15}{filename}\n"
            
            # 3. Tính chính xác dòng hiện tại (Header = 1, file đầu tiên = 2, ...)
            line_num = idx + 1
            
            # Chèn dòng và áp dụng màu nền
            text_widget.insert(tk.END, line, row_tag)

            # 4. Quét và Highlight từ khóa
            if keyword:
                words = keyword.split()
                filename_lower = filename.lower()
                prefix_len = len(f"{idx:<5}{tag_label:<15}")

                for word in words:
                    for match in re.finditer(re.escape(word), filename_lower):
                        real_start = prefix_len + match.start()
                        real_end = prefix_len + match.end()
                        
                        text_widget.tag_add(
                            "highlight",
                            f"{line_num}.{real_start}",
                            f"{line_num}.{real_end}"
                        )

        # 5. ÉP THẺ HIGHLIGHT NỔI LÊN TRÊN CÙNG ĐỂ KHÔNG BỊ MÀU NỀN CHE MẤT
        text_widget.tag_raise("highlight")
        text_widget.config(state="disabled")

    def handle_download(event):
        index = text_widget.index("@%d,%d" % (event.x, event.y))
        line = int(index.split(".")[0])

        if line <= 1:
            return

        file_index = line - 2
        if file_index < 0 or file_index >= len(filtered_files):
            return

        file = filtered_files[file_index]
        temp_path = download_file(token, file["id"], file["name"])

        if not temp_path:
            messagebox.showerror("Lỗi", f"Tải file {file['name']} thất bại!")
            return

        final_path = os.path.join(DOCUMENTARY_ARCHIVE_DIR, file["name"])
        try:
            shutil.move(temp_path, final_path)
        except:
            shutil.copy2(temp_path, final_path)
            os.remove(temp_path)

        update_table()

    root = tk.Tk()
    root.title("📁 RMC DRIVE VIEWER (OneDrive - Azure)")
    root.geometry("900x600")

    # ==== Thanh tìm kiếm ====
    frame_search = tk.Frame(root)
    frame_search.pack(pady=5, padx=5, fill="x")

    search_var = tk.StringVar()
    entry_search = tk.Entry(frame_search, textvariable=search_var, font=("Arial", 12), width=50)
    entry_search.pack(side="left", padx=5)

    btn_refresh = tk.Button(frame_search, text="🔄 Làm mới", font=("Arial", 12), command=update_table)
    btn_refresh.pack(side="right", padx=5)

    btn_open_folder = tk.Button(frame_search, text="📂", font=("Arial", 12),
                                command=lambda: os.startfile(DOCUMENTARY_ARCHIVE_DIR))
    btn_open_folder.pack(side="right", padx=5)

    # Ràng buộc sự kiện nhấn phím và nhấn Enter
    entry_search.bind("<KeyRelease>", update_table)
    entry_search.bind("<Return>", update_table)

    # ==== Bảng file ====
    frame_table = tk.Frame(root)
    frame_table.pack(pady=10, fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_table)
    scrollbar.pack(side="right", fill="y")

    text_widget = tk.Text(
        frame_table,
        wrap="none",
        font=("Consolas", 10),
        yscrollcommand=scrollbar.set
    )
    text_widget.pack(fill="both", expand=True)
    scrollbar.config(command=text_widget.yview)

    text_widget.bind("<Double-Button-1>", handle_download)

    # ==== Cấu hình màu sắc ====
    text_widget.tag_config(
        "highlight",
        background="yellow",
        foreground="black", # Đổi sang chữ đen trên nền vàng để dễ đọc hơn
        font=("Consolas", 10, "bold")
    )
    text_widget.tag_config("downloaded", background="#d0f0c0")
    text_widget.tag_config("not_downloaded", background="#f7c6c7")
    text_widget.tag_config("header", font=("Consolas", 10, "bold"))

    update_table()
    root.mainloop()
# == Cửa sổ tương tác dữ liệu ==
def create_data_interaction_window(root, title="Cửa sổ tương tác dữ liệu"):
    # =====================================================
    # KHỞI TẠO CỬA SỔ
    # =====================================================
    new_window = tk.Toplevel(root)
    new_window.title(title)
    new_window.geometry("600x420")
    new_window.configure(bg="white")
    new_window.transient(root)
    new_window.grab_set()

    current_files = []

    # =====================================================
    # DANH SÁCH THƯ MỤC HỆ THỐNG
    # =====================================================
    TARGET_FOLDERS = [AEONGMS_DIR, AEONMAXVALU_DIR, REPORT_FORM_DIR]

    # =====================================================
    # GIAO DIỆN CHÍNH CỦA CỬA SỔ
    # =====================================================
    tk.Label(new_window, text="Chọn thư mục làm việc (Folders):", font=("Arial", 11, "bold"), bg="white").pack(pady=(15, 5), padx=20, anchor="w")
    folder_combo = ttk.Combobox(new_window, values=TARGET_FOLDERS, font=("Arial", 10), state="readonly")
    folder_combo.pack(padx=20, fill="x", pady=5)
    folder_combo.set("--- Bấm vào đây để chọn thư mục ---")

    search_frame = tk.Frame(new_window, bg="white")
    search_frame.pack(padx=20, fill="x", pady=(10, 5))
    tk.Label(search_frame, text="Tìm tệp tin:", font=("Arial", 10, "bold"), bg="white").pack(side="left")
    search_entry = tk.Entry(search_frame, font=("Arial", 10), bd=1, relief="solid")
    search_entry.pack(side="left", padx=(10, 20), fill="x", expand=True)
    status_label = tk.Label(search_frame, text="Chưa chọn thư mục", font=("Arial", 10, "italic"), bg="white", fg="gray")
    status_label.pack(side="right")

    tk.Label(new_window, text="Danh sách các tệp tin (Files) trong thư mục:", font=("Arial", 11, "bold"), bg="white").pack(pady=(10, 5), padx=20, anchor="w")
    file_frame = tk.Frame(new_window, bg="white")
    file_frame.pack(padx=20, fill="x")
    file_scrollbar = tk.Scrollbar(file_frame, orient="vertical")
    file_scrollbar.pack(side="right", fill="y")
    file_listbox = tk.Listbox(file_frame, font=("Arial", 10), height=4, yscrollcommand=file_scrollbar.set)
    file_listbox.pack(side="left", fill="x", expand=True)
    file_scrollbar.config(command=file_listbox.yview)

    # Logic lọc file gốc
    def filter_files(*args):
        keyword = search_entry.get().strip().lower()
        file_listbox.delete(0, tk.END)
        if not current_files:
            return
        filtered_files = [f for f in current_files if keyword in f.lower()]
        if not filtered_files:
            file_listbox.insert(tk.END, "(Không tìm thấy tệp tin phù hợp)")
            status_label.config(text="Đã tìm thấy 0 file liên quan", fg="red")
        else:
            for f in filtered_files:
                file_listbox.insert(tk.END, f)
            status_label.config(text=f"Đã tìm thấy {len(filtered_files)} file liên quan", fg="#5A780B")

    def on_folder_combo_select(event):
        nonlocal current_files
        selected_folder = folder_combo.get()
        search_entry.delete(0, tk.END)
        file_listbox.delete(0, tk.END)
        current_files = []
        try:
            items = os.listdir(selected_folder)
            current_files = [f for f in items if os.path.isfile(os.path.join(selected_folder, f))]
            filter_files()
        except Exception:
            file_listbox.insert(tk.END, "(Lỗi đọc thư mục)")
            status_label.config(text="Lỗi", fg="red")

    folder_combo.bind("<<ComboboxSelected>>", on_folder_combo_select)
    search_entry.bind("<KeyRelease>", filter_files)

    # =====================================================
    # LOGIC MỚI: TÌM TỪ KHÓA CHUNG / KHÁC BIỆT GIỮA CÁC FILE
    # =====================================================
    def get_batch_diff_content(folder_path, file_names):
        if not file_names: 
            return "📌 KHÔNG CÓ FILE NÀO ĐƯỢC CHỌN.\n"
            
        files_lines = []
        try:
            for filename in file_names:
                filepath = os.path.join(folder_path, filename)
                # Đọc toàn bộ dòng của từng file
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    files_lines.append(f.read().splitlines())
        except Exception as e:
            return f"❌ Lỗi khi đọc file: {str(e)}"
                
        if not files_lines: return "📌 KHÔNG ĐỌC ĐƯỢC NỘI DUNG.\n"
            
        # Tìm số lượng dòng ngắn nhất và dài nhất trong tất cả các file
        min_lines = min(len(lines) for lines in files_lines)
        max_lines = max(len(lines) for lines in files_lines)
            
        result_lines = []
        
        # Biến đếm để tính toán tỉ lệ giống nhau
        total_matches = 0
        total_words = 0
        
        for i in range(min_lines):
            # Lấy dòng thứ i của tất cả các file
            lines_at_i = [f_lines[i] for f_lines in files_lines]
                
            # Tách dòng thành các từ (từ khóa) cách nhau bởi khoảng trắng
            words_list = [line.split() for line in lines_at_i]
                
            # Nếu tất cả các file đều là dòng trống
            if all(not w for w in words_list):
                result_lines.append("")
                continue
                    
            min_words = min(len(w) for w in words_list)
            max_words = max(len(w) for w in words_list)
            
            # Cộng dồn số từ nhiều nhất vào tổng số từ của dòng này
            total_words += max_words
                
            line_result = []
            for j in range(min_words):
                first_word = words_list[0][j]
                # Kiểm tra xem từ ở vị trí j có giống nhau ở tất cả các file không
                if all(w_list[j] == first_word for w_list in words_list):
                    line_result.append(first_word)
                    total_matches += 1  # Tăng biến đếm nếu từ khóa khớp hoàn toàn
                else:
                    # Gộp các [different keyword] liền kề lại cho đỡ rối mắt
                    if not line_result or line_result[-1] != "[different keyword]":
                        line_result.append("[different keyword]")
                
            # Nếu có file dài hơn (dư từ khóa ở cuối dòng)
            if max_words > min_words:
                if not line_result or line_result[-1] != "[different keyword]":
                    line_result.append("[different keyword]")
                        
            result_lines.append(" ".join(line_result))
                
        # Nếu có file nhiều dòng hơn file khác
        if max_lines > min_lines:
            # Cộng thêm trọng số khác biệt cho các dòng chênh lệch
            total_words += (max_lines - min_lines)
            result_lines.append("\n... [different keyword] (Các file có số lượng dòng khác nhau) ...")
            
        # -----------------------------------------------------------
        # LOGIC CẢNH BÁO: Kiểm tra cấu trúc và nội dung hoàn toàn khác nhau
        # Nếu tổng số từ > 0 và tỉ lệ giống nhau < 15% (0.15)
        # -----------------------------------------------------------
        if total_words > 0 and (total_matches / total_words) < 0.15:
            return "📌 Có nhiều file nội dung khác nhau, vui lòng lọc lại."
                
        return "\n".join(result_lines)

    # =====================================================
    # GIAO DIỆN UPDATE
    # =====================================================
    def open_update_window():
        filtered_files = file_listbox.get(0, tk.END)
        if not filtered_files or filtered_files[0].startswith("("):
            messagebox.showwarning("Cảnh báo", "Không có tệp tin hợp lệ nào để cập nhật!")
            return

        selected_folder = folder_combo.get()

        upd_window = tk.Toplevel(new_window)
        upd_window.title("Cập nhật nội dung File")
        upd_window.geometry("750x450")
        upd_window.configure(bg="#f4f4f4")
        upd_window.transient(new_window)
        upd_window.grab_set()

        # --- BÊN TRÁI: DANH SÁCH FILE ---
        left_frame = tk.Frame(upd_window, bg="white", bd=1, relief="solid")
        left_frame.pack(side="left", fill="y", padx=15, pady=15)
        tk.Label(left_frame, text="Danh sách file đã lọc", font=("Arial", 11, "bold"), bg="white").pack(pady=5)
        
        upd_listbox_frame = tk.Frame(left_frame)
        upd_listbox_frame.pack(fill="both", expand=True, padx=5, pady=5)
        upd_scroll = tk.Scrollbar(upd_listbox_frame, orient="vertical")
        upd_scroll.pack(side="right", fill="y")
        upd_listbox = tk.Listbox(upd_listbox_frame, width=35, yscrollcommand=upd_scroll.set, font=("Arial", 10))
        upd_listbox.pack(side="left", fill="both", expand=True)
        upd_scroll.config(command=upd_listbox.yview)

        for f in filtered_files:
            upd_listbox.insert(tk.END, f)

        # --- BÊN PHẢI: KHU VỰC THAO TÁC ---
        right_frame = tk.Frame(upd_window, bg="#f4f4f4")
        right_frame.pack(side="left", fill="both", expand=True, padx=(0, 15), pady=15)

        top_right_frame = tk.Frame(right_frame, bg="#f4f4f4")
        top_right_frame.pack(fill="x", pady=(0, 10))

        # --- COMBOBOX THAY THẾ CHO RADIOBUTTON ---
        radio_frame = tk.Frame(top_right_frame, bg="#f4f4f4")
        radio_frame.pack(side="left")
        
        tk.Label(radio_frame, text="Chế độ:", font=("Arial", 10, "bold"), bg="#f4f4f4").pack(side="left", padx=(0, 5))
        mode_combo = ttk.Combobox(radio_frame, values=["1 File", "Hàng loạt"], state="readonly", width=12, font=("Arial", 10))
        mode_combo.set("1 File")
        mode_combo.pack(side="left")

        btn_action_frame = tk.Frame(top_right_frame, bg="#f4f4f4")
        btn_action_frame.pack(side="right", anchor="n")
        tk.Button(btn_action_frame, text="Sửa", width=10, bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=5)
        tk.Button(btn_action_frame, text="Xóa", width=10, bg="#F44336", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=5)

        content_frame = tk.LabelFrame(right_frame, text=" Hiển thị nội dung trong file được chọn ", font=("Arial", 10, "bold"), bg="white", bd=1, relief="solid")
        content_frame.pack(fill="both", expand=True)

        content_text = tk.Text(content_frame, font=("Consolas", 10), wrap="word")
        content_text.pack(fill="both", expand=True, padx=5, pady=5)

        # --- LOGIC ĐIỀU KHIỂN CHẾ ĐỘ ---
        def handle_mode_change(*args):
            mode = mode_combo.get()
            content_text.config(state=tk.NORMAL)
            content_text.delete("1.0", tk.END)

            if mode == "1 File":
                content_text.insert(tk.END, "👉 Hãy chọn một file bên trái để xem nội dung...")
                content_text.config(state=tk.DISABLED) 

            elif mode == "Hàng loạt":
                upd_listbox.selection_clear(0, tk.END)
                
                content_text.insert(tk.END, "⏳ Đang quét dữ liệu, phân tích khác biệt...\n\n")
                upd_window.update() 
                
                content_text.delete("1.0", tk.END)
                common_text = get_batch_diff_content(selected_folder, filtered_files)
                content_text.insert(tk.END, "--- NỘI DUNG TỔNG HỢP (HÀNG LOẠT) ---\n\n" + common_text)
                
                content_text.config(state=tk.DISABLED)

        # Gắn sự kiện khi đổi giá trị Combobox
        mode_combo.bind("<<ComboboxSelected>>", handle_mode_change)

        # Bắt sự kiện khi click vào một tệp tin trong danh sách ở chế độ 1 File
        def on_upd_listbox_select(event):
            if mode_combo.get() == "1 File":
                selection = upd_listbox.curselection()
                if not selection: return
                
                selected_filename = upd_listbox.get(selection[0])
                filepath = os.path.join(selected_folder, selected_filename)
                
                content_text.config(state=tk.NORMAL)
                content_text.delete("1.0", tk.END)
                
                try:
                    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                        file_content = f.read()
                        content_text.insert(tk.END, file_content)
                except Exception as e:
                    content_text.insert(tk.END, f"❌ Lỗi khi mở file: {str(e)}")

        upd_listbox.bind("<<ListboxSelect>>", on_upd_listbox_select)
        
        # Khởi chạy logic lúc mới bật form Update
        handle_mode_change()

    # =====================================================
    # NÚT ĐIỀU KHIỂN MAIN WINDOW
    # =====================================================
    control_frame = tk.Frame(new_window, bg="white")
    control_frame.pack(pady=20)
    tk.Button(control_frame, text="Update", font=("Arial", 11, "bold"), bg="#4a90e2", fg="white", width=12, command=open_update_window).pack(side="left", padx=10)
    tk.Button(control_frame, text="Đóng", font=("Arial", 11), bg="#ff4c4c", fg="white", width=12, command=new_window.destroy).pack(side="left", padx=10)

# === Khu vực tạo các thành phần =======================================================================================================================================
copy_button = tk.Button(left_controls, text="Copy", font=("Arial", 10, "bold"), bg="#4CAF50", fg="white",
                        command=copy_text_to_clipboard, width=15)
copy_button.pack(side="left", padx=(0, 5))

clear_button = tk.Button(left_controls, text="Clear", font=("Arial", 10, "bold"), bg="#f44336", fg="white",
                         command=clear_text_output, width=15)
clear_button.pack(side="left")

# Catch (ngoài cùng bên phải)
catch_button = tk.Button(right_controls, text="Catch", font=("Arial", 10, "bold"), bg="#029B82", fg="white",
                         command=catch_clock, width=10)
catch_button.pack(side='right', padx=5)

# Clock (giữa)
clock_label = tk.Label(right_controls, font=("Roboto", 20, "bold"), fg="#D20103",)
clock_label.pack(side='right', padx=10)

# Continue (ngoài cùng bên trái của cụm)
continue_button = tk.Button(right_controls, text="Continue", font=("Arial", 10, "bold"), bg="#2196F3", fg="white",
                            command=continue_clock, width=10)
continue_button.pack(side='right', padx=5)

# ==== NÚT CONTACT ====
def contact_action():
    if fill_box(1):  # Chỉ tô nếu ô 0 đã tô
        create_new_window_contact("Contact")
        on_category_click()
        reset_timer()
contact_button = tk.Button(left_button_frame, text="Contact", font=("Arial", 12, "bold"),
                           bg="#2196F3", fg="white", width=10, command=lambda: contact_action())
contact_button.pack(pady=5)

# ==== NÚT STATUS ====
def status_action():
    if fill_box(2):  # Chỉ tô nếu ô 1 đã tô
        create_new_window_status("Status")
        on_category_click()
        reset_timer()
status_button = tk.Button(left_button_frame, text="Status", font=("Arial", 12, "bold"),
                          bg="#FF9800", fg="white", width=10, command=lambda: status_action())
status_button.pack(pady=5)

# ==== NÚT NOTE ====
def note_action():
    create_new_window_note()
note_button = tk.Button(left_button_frame, text="Note", font=("Arial", 12, "bold"),
                        bg="#873e23", fg="white", width=10, command=lambda: note_action())
note_button.pack(pady=5)

# ==== NÚT KHO ẢNH DAVITEQ ====
def image_daviteq_action():
    create_new_window_image_daviteq("IMAGE")
image_daviteq_button = tk.Button(left_button_frame, text="IMAGE", font=("Arial", 12, "bold"),
                                 bg="#3fc4f3", fg="white", width=10, command=lambda: image_daviteq_action())
image_daviteq_button.pack(pady=5)

# ==== NÚT vào KHO DOCUMENTARY ====
def rmc_drive_viewer_action():
    create_documentary_viewer(access_token, documentary_archive_url)

rmc_drive_viewer_button = tk.Button(
    left_button_frame,
    text="Document",
    font=("Arial", 12, "bold"),bg="#5A780B", fg="white",width=10,
    command=rmc_drive_viewer_action
)
rmc_drive_viewer_button.pack(pady=5)

# ==== NÚT vào DATA INTERACT ====
def rmc_date_interact_action():
    # Phải truyền biến 'root' vào để hàm nhận diện được cửa sổ chính
    create_data_interaction_window(root)

rmc_date_interact_button = tk.Button(
    left_button_frame,
    text="DATA",
    font=("Arial", 12, "bold"),
    bg="#9d00ff",
    fg="white",
    width=10,
    command=rmc_date_interact_action
)
rmc_date_interact_button.pack(pady=5)

# ==== NÚT XÁC NHẬN HÀNH ĐỘNG ====
def confirm_action():
    for i in range(3, 6):
        if not box_filled[i]:
            if fill_box(i):
                on_category_click()
                break  # Đảm bảo chỉ tô một ô
            else:
                break  # Dừng lại nếu chưa đủ điều kiện
confirm_button = tk.Button(main_frame, text="Xác nhận", font=("Arial", 12, "bold"),
                           bg="#4CAF50", fg="white", command=confirm_action)
confirm_button.pack(pady=10)
# =========================================================
# CLOSE APP SAFELY
# =========================================================
def show_startup_window():
    # =====================================================
    # FIRST RUN
    # =====================================================
    if is_first_run():
        root.withdraw()
        startup = tk.Toplevel(root)
        startup.title("RMC Assistant")
        startup.geometry("450x220")
        startup.resizable(False, False)
        startup.grab_set()
        tk.Label(
            startup,
            text="🚀 Thiết lập lần đầu",
            font=("Arial", 18, "bold")
        ).pack(pady=(25, 15))

        status_label = tk.Label(
            startup,
            text="Đang tải dữ liệu lần đầu...",
            font=("Arial", 11),
            fg="blue"
        )

        status_label.pack(pady=10)
        progress = ttk.Progressbar(
            startup,
            mode="indeterminate",
length=300
        )

        progress.pack(pady=15)
        progress.start()
        # =====================================================
        # THREAD FIRST SETUP
        # =====================================================
        def run_first_setup():
            try:
                # =========================
                # LOGIN
                # =========================
                root.after(
                    0,
                    lambda: status_label.config(
                        text="🔑 Đang đăng nhập Azure..."
                    )
                )
                graph_session.ensure_token()
                # =========================
                # BUILD DATA LINK
                # =========================
                root.after(
                    0,
                    lambda: status_label.config(
                        text="🔄 Đang build data link..."
                    )
                )
                build_data_link()
                # =========================
                # AUTO SYNC
                # =========================
                root.after(
                    0,
                    lambda: status_label.config(
                        text="🔄 Đang đồng bộ dữ liệu..."
                    )
                )
                auto_sync_all_onedrive()
                # =========================
                # CREATE FIRST RUN FILE
                # =========================
                create_first_run_flag()
                root.after(
                    0,
                    lambda: status_label.config(
                        text="✅ Thiết lập hoàn tất"
                    )
                )
                # =====================================================
                # LOAD CACHE + OPEN MAIN APP
                # =====================================================
                def finish_setup():
                        try:
                            startup.destroy()
                        except:
                            pass
                        global DATA_LINK
                        DATA_LINK = load_data_link_json()
                        # =========================================
                        # BUILD DEVICE MAP LẠI
                        # =========================================
                        refresh_report_mapping()
                        # =========================================
                        # HIỆN GIAO DIỆN CHỌN
                        # =========================================
                        show_startup_window()
                root.after(0, finish_setup)   
            except Exception as e:
                print("❌ FIRST RUN ERROR:", e)
                def show_error():
                    messagebox.showerror(
                        "First Run Error",
                        str(e)
                    )
                    try:
                        startup.destroy()
                    except:
                        pass
                root.after(0, show_error)
        threading.Thread(
            target=run_first_setup,
            daemon=True
        ).start()
        return
    # =====================================================
    # NORMAL STARTUP
    # =====================================================
    root.withdraw()
    startup = tk.Toplevel(root)
    startup.title("RMC Assistant")
    startup.geometry("450x250")
    startup.resizable(False, False)
    startup.grab_set()
    # =========================
    # TITLE
    # =========================
    tk.Label(
        startup,
        text="RMC Assistant",
        font=("Arial", 18, "bold")
    ).pack(pady=(25, 10))
    # =========================
    # STATUS
    # =========================
    status_label = tk.Label(
        startup,
        text="Chọn chế độ khởi động",
        font=("Arial", 11),fg="blue"
    )
    status_label.pack(pady=(0, 20))
    # =========================
    # BUTTON FRAME
    # =========================
    btn_frame = tk.Frame(startup)
    btn_frame.pack(pady=10)
    # =========================
    # OFFLINE
    # =========================
    tk.Button(
        btn_frame,
        text="▶ Vào chương trình",
        font=("Arial", 11, "bold"),
        width=25,height=2,bg="#4CAF50",fg="white",
        command=lambda: skip_sync(startup)
    ).pack(pady=10)
    # =========================
    # ONLINE SYNC
    # =========================
    tk.Button(
        btn_frame,
        text="🔄 Đồng bộ dữ liệu",
        font=("Arial", 11, "bold"),
        width=25,height=2,bg="#2196F3",fg="white",
        command=lambda:
            threading.Thread(
                target=run_sync_and_launch,
                args=(startup, status_label),
                daemon=True
            ).start()
    ).pack(pady=10)
# ==== CHẠY Giao diện hỏi người dùng có muốn đồng bộ hay ko đồng bọ trước khi vào app chính ====
show_startup_window()
# ==== CHẠY ỨNG DỤNG ====
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()

