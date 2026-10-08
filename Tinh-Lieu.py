import io
import json
import math
import os
import re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import openpyxl
from openpyxl import load_workbook
import streamlit as st

# --- CẤU HÌNH GOOGLE DRIVE ---
FOLDER_ID = "1GbnN63XfIc1UmR_2XW8LPePpFxvBB__e"
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
EXCEL_FILE_NAME = "danh_sach_thuoc_lieu_dung.xlsx"
# 1. Nhúng trực tiếp cấu hình OAuth Credentials vào code
CLIENT_CONFIG = {
    "installed": {
        "client_id": "976291028498-mkqsg7fs123gt1hjs8kh23epmqklq31s.apps.googleusercontent.com",
        "project_id": "data-508701",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_secret": "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",
        "redirect_uris": ["http://localhost"]
    }
}
@st.cache_resource
def get_drive_service():
    # Tự động lấy thông tin từ Streamlit Secrets
    client_id = st.secrets["client_id"]
    client_secret = st.secrets["client_secret"]
    refresh_token = st.secrets["refresh_token"]

    # Khởi tạo credentials từ Refresh Token
    creds = Credentials(
        token=None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=SCOPES
    )

    # Làm mới Token để có Access Token hợp lệ
    creds.refresh(Request())
    
    return build('drive', 'v3', credentials=creds)

# --- Streamlit UI ---
st.title("Google Drive Streamlit App")

try:
    service = get_drive_service()
    st.success("✅ Kết nối Google Drive thành công!")

    # Chạy thử: Lấy danh sách 10 file đầu tiên trên Drive
    results = service.files().list(
        pageSize=10, 
        fields="nextPageToken, files(id, name)"
    ).execute()
    items = results.get('files', [])

    st.subheader("Danh sách file gần đây:")
    if not items:
        st.write("Không tìm thấy file nào.")
    else:
        for item in items:
            st.write(f"📄 **{item['name']}** (`{item['id']}`)")

except Exception as e:
    st.error(f"❌ Lỗi kết nối Google Drive: {e}")
def search_file_in_folder(service, filename, folder_id):
    """Tìm file theo tên trong FOLDER_ID cụ thể."""
    query = f"'{folder_id}' in parents and name = '{filename}' and trashed = false"
    results = (
        service.files()
        .list(q=query, spaces="drive", fields="files(id, name)")
        .execute()
    )
    items = results.get("files", [])
    if items:
        return items[0]["id"]
    return None


def upload_pdf_to_drive(uploaded_file):
    try:
        service = get_drive_service()
        if not service:
            st.error("❌ Không tìm thấy file credentials.json để xác thực Drive.")
            return None, None

        file_metadata = {
            "name": uploaded_file.name,
            "parents": [FOLDER_ID],
        }
        media = MediaIoBaseUpload(
            io.BytesIO(uploaded_file.getvalue()),
            mimetype="application/pdf",
            resumable=True,
        )
        file = (
            service.files()
            .create(
                body=file_metadata,
                media_body=media,
                fields="id, webViewLink",
            )
            .execute()
        )
        return file.get("id"), file.get("webViewLink")
    except Exception as e:
        st.error(f"Lỗi Upload PDF lên Drive: {e}")
        return None, None


# --- CẤU HÌNH TRANG WEB ---
st.set_page_config(
    page_title="Công Cụ Tính Liều Thuốc - Minh Nhân Professional",
    page_icon="🩺",
    layout="wide",
)

st.title("🩺 Công Cụ Tính Liều Thuốc & Bút Tiêm Insulin")
st.markdown("---")


# --- HÀM ĐỌC DỮ LIỆU EXCEL TỪ GOOGLE DRIVE HOẶC FILE UPLOAD ---
@st.cache_data(show_spinner=False)
def load_excel_data_from_drive_or_file(uploaded_file_bytes=None):
    try:
        file_stream = None

        if uploaded_file_bytes is not None:
            file_stream = io.BytesIO(uploaded_file_bytes)
        else:
            service = get_drive_service()
            if service:
                file_id = search_file_in_folder(
                    service, EXCEL_FILE_NAME, FOLDER_ID
                )
                if file_id:
                    request = service.files().get_media(fileId=file_id)
                    file_stream = io.BytesIO()
                    downloader = MediaIoBaseDownload(file_stream, request)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()
                    file_stream.seek(0)

        if file_stream is None:
            return {}, {}, {}, {}, {}

        wb = load_workbook(file_stream)
        ws = wb.active

        thuoc_dict = {}
        thuoc_quycach_dict = {}
        thuoc_chidinh_dict = {}
        thuoc_chongchidinh_dict = {}
        thuoc_bietduoc_dict = {}

        for row in ws.iter_rows(min_row=1, values_only=True):
            if row and row[0]:  # Cột A: Tên thuốc / Hoạt chất
                key = str(row[0]).strip()

                # Cột B: Liều dùng
                thuoc_dict[key] = (
                    str(row[1]).strip() if len(row) > 1 and row[1] else ""
                )

                # Cột C: Quy cách
                thuoc_quycach_dict[key] = (
                    str(row[2]).strip() if len(row) > 2 and row[2] else ""
                )

                # Cột D: Chỉ định
                thuoc_chidinh_dict[key] = (
                    str(row[3]).strip() if len(row) > 3 and row[3] else ""
                )

                # Cột E: Chống chỉ định
                thuoc_chongchidinh_dict[key] = (
                    str(row[4]).strip() if len(row) > 4 and row[4] else ""
                )

                # Cột F: Biệt dược (Dạng JSON string)
                biet_duoc_raw = (
                    str(row[5]).strip() if len(row) > 5 and row[5] else ""
                )
                bd_list = []
                if biet_duoc_raw:
                    try:
                        bd_list = json.loads(biet_duoc_raw)
                    except Exception:
                        bd_list = [{"ten": biet_duoc_raw, "link": ""}]
                thuoc_bietduoc_dict[key] = bd_list

        return (
            thuoc_dict,
            thuoc_quycach_dict,
            thuoc_chidinh_dict,
            thuoc_chongchidinh_dict,
            thuoc_bietduoc_dict,
        )
    except Exception as e:
        st.error(f"❌ Lỗi đọc dữ liệu Excel: {e}")
        return {}, {}, {}, {}, {}


# --- KHU VỰC QUẢN LÝ FILE EXCEL ---
st.sidebar.header("⚙️ Quản lý Dữ liệu")
uploaded_file = st.sidebar.file_uploader(
    "Tải lên file Excel danh sách thuốc (.xlsx)", type=["xlsx"]
)

if uploaded_file is None:
    (
        thuoc_dict,
        thuoc_quycach_dict,
        thuoc_chidinh_dict,
        thuoc_chongchidinh_dict,
        thuoc_bietduoc_dict,
    ) = load_excel_data_from_drive_or_file()
    if thuoc_dict:
        st.sidebar.success("✅ Đang sử dụng dữ liệu Excel từ Google Drive.")
    else:
        st.sidebar.warning(
            "⚠️ Chưa tìm thấy file Excel trên Drive hoặc chưa kết nối."
        )
else:
    (
        thuoc_dict,
        thuoc_quycach_dict,
        thuoc_chidinh_dict,
        thuoc_chongchidinh_dict,
        thuoc_bietduoc_dict,
    ) = load_excel_data_from_drive_or_file(uploaded_file.getvalue())
    st.sidebar.success("🎉 Đã cập nhật dữ liệu từ file đính kèm!")

danh_sach_thuoc = sorted(thuoc_dict.keys())

# --- MENU CHỨC NĂNG ---
if "menu_selection" not in st.session_state:
    st.session_state["menu_selection"] = "⚖️ Tính Liều Theo Cân Nặng"

st.markdown(
    """
    <style>
    div.stRadio > label { font-size: 24px !important; font-weight: bold !important; color: #1f77b4; padding-bottom: 12px; }
    div[role="radiogroup"] span[data-baseweb="radio"] { display: none !important; }
    div[role="radiogroup"] label p { font-size: 20px !important; font-weight: bold !important; margin: 0px !important; }
    div[role="radiogroup"] label {
        background-color: #f0f2f6; padding: 12px 25px !important;
        border-radius: 10px; margin-right: 10px; box-shadow: 2px 2px 6px rgba(0,0,0,0.08);
        cursor: pointer; transition: all 0.2s ease; border: 2px solid transparent;
    }
    div[role="radiogroup"] label:hover { background-color: #e6f2ff; transform: translateY(-2px); border: 2px solid #1f77b4; }
    </style>
    """,
    unsafe_allow_html=True,
)

chon_tab = st.radio(
    "📌 Chọn chức năng:",
    options=[
        "⚖️ Tính Liều Theo Cân Nặng",
        "🖊️ Tính Số Lượng Bút Insulin Theo Ngày Kê",
        "➕ Thêm/Cập Nhật Dữ Liệu Thuốc",
    ],
    key="menu_selection",
    horizontal=True,
)

st.markdown("<br>", unsafe_allow_html=True)


# ==============================================================================
# CHỨC NĂNG 1: TÍNH LIỀU THEO CÂN NẶNG
# ==============================================================================
if chon_tab == "⚖️ Tính Liều Theo Cân Nặng":
    col1, col2 = st.columns([2, 3])

    with col1:
        st.subheader("Nhập thông tin")
        ten_thuoc_chon = st.selectbox(
            "1. Nhập/Chọn tên thuốc (Hoạt chất):",
            options=[""] + danh_sach_thuoc,
            key="tab1_thuoc",
        )

        cac_lieu_dung = []
        if ten_thuoc_chon:
            value_b = thuoc_dict.get(ten_thuoc_chon, "")
            if value_b:
                cac_lieu_dung = [
                    v.strip() for v in value_b.split(";") if v.strip()
                ]

        lieu_chon = st.selectbox(
            "2. Chọn liều dùng tương ứng:",
            options=(
                cac_lieu_dung
                if cac_lieu_dung
                else ["(Vui lòng chọn thuốc trước)"]
            ),
        )

        can_nang = st.number_input(
            "3. Nhập cân nặng bệnh nhân (kg):",
            min_value=0.0,
            max_value=200.0,
            value=0.0,
            step=0.1,
            format="%.1f",
        )

    with col2:
        st.subheader("📋 Kết quả đề nghị")

        if ten_thuoc_chon and can_nang > 0 and lieu_chon:
            numbers = re.findall(r"[\d.]+", lieu_chon)
            if numbers:
                lieu_dung_so = float(numbers[0])
                tong_lieu = lieu_dung_so * can_nang
                is_lan_dung = "ngày" in lieu_chon.lower()
                tong_lieu_target = (
                    tong_lieu / 2 if is_lan_dung else tong_lieu
                )
                lieu_theo_tuoi = "tuổi" in lieu_chon.lower()

                if not lieu_theo_tuoi:
                    st.metric(
                        label="Tổng liều tính toán:",
                        value=f"{tong_lieu:.0f} mg",
                    )
                else:
                    st.info("💡 Lưu ý: Xem liều dùng chi tiết theo tuổi.")

                quy_cach = thuoc_quycach_dict.get(ten_thuoc_chon, "")
                quy_cach_base = []
                if quy_cach:
                    quy_cach_base = [
                        float(x.strip())
                        for x in quy_cach.split(";")
                        if x.strip().replace(".", "", 1).isdigit()
                    ]

                if quy_cach_base:
                    phuong_an_list = []
                    for qc in quy_cach_base:
                        phuong_an_list.append(
                            {"qc": qc, "he_so": 0.5, "mg": qc * 0.5}
                        )
                        phuong_an_list.append(
                            {"qc": qc, "he_so": 1.0, "mg": qc * 1.0}
                        )
                        phuong_an_list.append(
                            {"qc": qc, "he_so": 2.0, "mg": qc * 2.0}
                        )

                    priority_map = {1.0: 1, 0.5: 2, 2.0: 3}
                    best = min(
                        phuong_an_list,
                        key=lambda p: (
                            abs(p["mg"] - tong_lieu_target),
                            priority_map[p["he_so"]],
                        ),
                    )

                    qc_mg = (
                        int(best["qc"])
                        if best["qc"].is_integer()
                        else best["qc"]
                    )
                    he_so = best["he_so"]

                    if not lieu_theo_tuoi:
                        if is_lan_dung:
                            desc = {
                                0.5: "Sáng 1/2 - Chiều 1/2",
                                1.0: "Sáng 1 - Chiều 1",
                                2.0: "Sáng 2 - Chiều 2",
                            }
                            st.success(
                                f"👉 **Đề nghị:** Dùng **{desc[he_so]}** loại **{qc_mg} mg**"
                            )
                        else:
                            desc = {
                                0.5: "1/2 viên/gói",
                                1.0: "1 viên/gói",
                                2.0: "2 viên/gói",
                            }
                            st.success(
                                f"👉 **Đề nghị:** Một lần dùng **{desc[he_so]}** loại **{qc_mg} mg**"
                            )
                    else:
                        st.warning(f"👉 **Đề nghị:** {lieu_chon}")
                else:
                    st.warning(
                        "⚠️ Chưa có cấu hình quy cách (Cột C) trong file Excel."
                    )
        else:
            st.info("Vui lòng chọn thuốc, liều và nhập cân nặng > 0.")

        st.markdown("---")

        # THÔNG TIN CHỈ ĐỊNH, CHỐNG CHỈ ĐỊNH VÀ BIỆT DƯỢC
        if ten_thuoc_chon:
            chi_dinh = thuoc_chidinh_dict.get(ten_thuoc_chon, "")
            chong_chi_dinh = thuoc_chongchidinh_dict.get(ten_thuoc_chon, "")
            list_bd = thuoc_bietduoc_dict.get(ten_thuoc_chon, [])

            with st.expander(
                "ℹ️ **Thông tin Chi tiết & Biệt dược đính kèm**", expanded=True
            ):
                st.markdown(f"✅ **Chỉ định:** {chi_dinh or 'Chưa có dữ liệu'}")
                st.markdown(
                    f"🚫 **Chống chỉ định:** :red[{chong_chi_dinh or 'Chưa có dữ liệu'}]"
                )

                st.markdown("💊 **Danh sách Biệt dược & Minh chứng PDF:**")
                if list_bd:
                    for bd in list_bd:
                        ten_bd = bd.get("ten", "Không rõ tên")
                        link_bd = bd.get("link", "")
                        if link_bd:
                            st.write(
                                f"- **{ten_bd}**: [📄 Xem file PDF minh chứng]({link_bd})"
                            )
                        else:
                            st.write(
                                f"- **{ten_bd}**: *(Chưa đính kèm PDF)*"
                            )
                else:
                    st.caption("Chưa có danh sách biệt dược.")


# ==============================================================================
# CHỨC NĂNG 2: TÍNH SỐ LƯỢNG BÚT INSULIN
# ==============================================================================
elif chon_tab == "🖊️ Tính Số Lượng Bút Insulin Theo Ngày Kê":
    col_in1, col_in2 = st.columns([2, 3])

    with col_in1:
        st.subheader("Nhập thông tin kê đơn")
        so_ngay_muon_ke = st.number_input(
            "Nhập số ngày muốn kê đơn (ngày):",
            min_value=1,
            max_value=365,
            value=31,
            step=1,
        )
        tong_lieu_mot_cay_but = st.number_input(
            "Tổng số liều của 1 cây bút tiêm (đơn vị):", value=300, step=50
        )
        lieu_sang = st.number_input(
            "Liều dùng buổi SÁNG (đơn vị):", min_value=0, value=0, step=1
        )
        lieu_chieu = st.number_input(
            "Liều dùng buổi CHIỀU (đơn vị):", min_value=0, value=0, step=1
        )

    with col_in2:
        st.subheader("📆 Kế hoạch sử dụng bút tiêm")
        tong_lieu_ngay = lieu_sang + lieu_chieu

        if tong_lieu_ngay > 0:
            tong_lieu_can_thiet = so_ngay_muon_ke * tong_lieu_ngay
            so_cay_but_tinh_duoc = (
                tong_lieu_can_thiet / tong_lieu_mot_cay_but
            )

            if so_cay_but_tinh_duoc.is_integer():
                so_cay_but_dieu_chinh = so_cay_but_tinh_duoc
            else:
                if so_ngay_muon_ke >= 90:
                    so_cay_but_dieu_chinh = max(
                        1, math.floor(so_cay_but_tinh_duoc)
                    )
                else:
                    so_cay_but_dieu_chinh = math.ceil(so_cay_but_tinh_duoc)

            so_ngay_dung = math.floor(
                (so_cay_but_dieu_chinh * tong_lieu_mot_cay_but) / tong_lieu_ngay
            )

            ket_qua_text = (
                f"🩺🖊️ Số cây bút cần kê: {math.floor(so_cay_but_dieu_chinh)} bút\n"
                f"📆 Số ngày dùng thực tế: {so_ngay_dung} ngày (Dự kiến kê: {so_ngay_muon_ke} ngày)\n"
                f"Các thuốc kèm theo trong đơn nếu có\n"
                f"💊💊 Sáng 1 Chiều 1: {so_ngay_dung * 2} viên/gói\n"
                f"💊💊💊💊 Sáng 2 Chiều 2: {so_ngay_dung * 4} viên/gói\n"
                f"💊💊💊 Sáng 1 Trưa 1 Chiều 1: {so_ngay_dung * 3} viên/gói\n"
                f"💊💊💊💊💊💊 Sáng 2 Trưa 2 Chiều 2: {so_ngay_dung * 6} viên/gói"
            )

            st.text_area(
                label="Kết quả chi tiết (Có thể sao chép để in):",
                value=ket_qua_text,
                height=220,
            )
        else:
            st.info("Vui lòng nhập liều tiêm sáng hoặc chiều để tính toán.")


# ==============================================================================
# CHỨC NĂNG 3: TRANG THÊM HOẠT CHẤT & BIỆT DƯỢC
# ==============================================================================
elif chon_tab == "➕ Thêm/Cập Nhật Dữ Liệu Thuốc":
    st.subheader("➕ Thêm mới / Cập nhật Thuốc & Biệt dược")
    st.info(
        "Nhập thông tin chi tiết dưới đây. Lưu ý các quy định phân cách bằng dấu chấm phẩy `;`"
    )

    hoat_chat = st.text_input(
        "1. Tên Hoạt chất (Cột A):", placeholder="Ví dụ: Paracetamol"
    )
    lieu_dung_input = st.text_input(
        "2. Quy định liều dùng (Cột B - cách nhau bởi dấu ';'):",
        placeholder="Ví dụ: 10 mg/kg/ngày; 15 mg/kg/ngày",
    )
    quy_cach_input = st.text_input(
        "3. Quy cách hàm lượng (Cột C - cách nhau bởi dấu ';'):",
        placeholder="Ví dụ: 80; 150; 250; 500",
    )
    chi_dinh_input = st.text_area("4. Chỉ định (Cột D):", height=80)
    chong_chi_dinh_input = st.text_area("5. Chống chỉ định (Cột E):", height=80)

    st.markdown("---")
    st.write("### 💊 Danh sách Biệt dược & File đính kèm (Cột F)")

    # Khởi tạo danh sách biệt dược trong Session State
    if "biet_duoc_list" not in st.session_state:
        st.session_state["biet_duoc_list"] = [{"id": 0}]

    def add_biet_duoc():
        st.session_state["biet_duoc_list"].append(
            {"id": len(st.session_state["biet_duoc_list"])}
        )

    for idx, item in enumerate(st.session_state["biet_duoc_list"]):
        col_bd1, col_bd2 = st.columns([1, 1])
        with col_bd1:
            st.text_input(
                f"Tên biệt dược #{idx+1}:", key=f"ten_bd_{item['id']}"
            )
        with col_bd2:
            st.file_uploader(
                f"PDF minh chứng #{idx+1}:",
                type=["pdf"],
                key=f"file_bd_{item['id']}",
            )

    st.button("➕ Thêm dòng biệt dược tiếp theo", on_click=add_biet_duoc)

    st.markdown("---")

    if st.button("💾 ĐỒNG BỘ VÀ LƯU VÀO GOOGLE DRIVE", type="primary"):
        if not hoat_chat.strip():
            st.error("❌ Vui lòng nhập Tên hoạt chất!")
        else:
            with st.spinner("Đang lưu dữ liệu và Upload lên Google Drive..."):
                biet_duoc_data = []

                # Xử lý upload các file PDF biệt dược
                for item in st.session_state["biet_duoc_list"]:
                    t_bd = st.session_state.get(f"ten_bd_{item['id']}", "")
                    f_bd = st.session_state.get(f"file_bd_{item['id']}", None)

                    if t_bd.strip():
                        pdf_link = ""
                        if f_bd is not None:
                            _, pdf_link = upload_pdf_to_drive(f_bd)

                        biet_duoc_data.append(
                            {"ten": t_bd.strip(), "link": pdf_link or ""}
                        )

                # Ghi và Đồng bộ File Excel trực tiếp trên Google Drive
                try:
                    service = get_drive_service()
                    if not service:
                        st.error(
                            "❌ Không kết nối được với Drive service. Vui lòng kiểm tra credentials.json/token.json"
                        )
                    else:
                        file_id = search_file_in_folder(
                            service, EXCEL_FILE_NAME, FOLDER_ID
                        )

                        if file_id:
                            # Đã có file trên Drive -> Tải nội dung về Memory
                            request = service.files().get_media(fileId=file_id)
                            file_stream = io.BytesIO()
                            downloader = MediaIoBaseDownload(
                                file_stream, request
                            )
                            done = False
                            while not done:
                                _, done = downloader.next_chunk()
                            file_stream.seek(0)
                            wb = openpyxl.load_workbook(file_stream)
                        else:
                            # Chưa có file -> Tạo workbook mới
                            wb = openpyxl.Workbook()

                        ws = wb.active

                        # Tìm xem Hoạt chất đã tồn tại chưa
                        target_row = None
                        for row_idx in range(1, ws.max_row + 1):
                            val_a = ws.cell(row=row_idx, column=1).value
                            if (
                                val_a
                                and str(val_a).strip().lower()
                                == hoat_chat.strip().lower()
                            ):
                                target_row = row_idx
                                break

                        if target_row is None:
                            target_row = (
                                ws.max_row + 1
                                if ws.cell(row=1, column=1).value
                                else 1
                            )

                        # Ghi thông tin vào hàng Excel
                        ws.cell(
                            row=target_row, column=1, value=hoat_chat.strip()
                        )  # Cột A
                        ws.cell(
                            row=target_row,
                            column=2,
                            value=lieu_dung_input.strip(),
                        )  # Cột B
                        ws.cell(
                            row=target_row,
                            column=3,
                            value=quy_cach_input.strip(),
                        )  # Cột C
                        ws.cell(
                            row=target_row,
                            column=4,
                            value=chi_dinh_input.strip(),
                        )  # Cột D
                        ws.cell(
                            row=target_row,
                            column=5,
                            value=chong_chi_dinh_input.strip(),
                        )  # Cột E
                        ws.cell(
                            row=target_row,
                            column=6,
                            value=json.dumps(
                                biet_duoc_data, ensure_ascii=False
                            ),
                        )  # Cột F

                        # Lưu Workbook ra BytesIO Stream
                        output_stream = io.BytesIO()
                        wb.save(output_stream)
                        output_stream.seek(0)

                        media = MediaIoBaseUpload(
                            output_stream,
                            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            resumable=True,
                        )

                        if file_id:
                            # Update đè file đã tồn tại trên Drive
                            service.files().update(
                                fileId=file_id, media_body=media
                            ).execute()
                        else:
                            # Tạo file mới trong FOLDER_ID trên Drive
                            file_metadata = {
                                "name": EXCEL_FILE_NAME,
                                "parents": [FOLDER_ID],
                            }
                            service.files().create(
                                body=file_metadata, media_body=media
                            ).execute()

                        st.cache_data.clear()  # Xóa cache Streamlit để tự động load dữ liệu mới nhất
                        st.success(
                            f"🎉 Đã đồng bộ thành công dữ liệu thuốc **{hoat_chat}** lên file Excel trên Google Drive!"
                        )

                        # Reset danh sách biệt dược về mặc định
                        st.session_state["biet_duoc_list"] = [{"id": 0}]

                except Exception as e:
                    st.error(f"❌ Lỗi đồng bộ lên Google Drive: {e}")
