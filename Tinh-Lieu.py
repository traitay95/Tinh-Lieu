import io
import json
import math
import re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import openpyxl
from openpyxl import load_workbook
import streamlit as st

# --- CẤU HÌNH GOOGLE DRIVE ---
FOLDER_ID = "1GbnN63XfIc1UmR_2XW8LPePpFxvBB__e"
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
EXCEL_FILE_NAME = "danh_sach_thuoc_lieu_dung.xlsx"


@st.cache_resource
def get_drive_service():
    client_id = st.secrets["client_id"]
    client_secret = st.secrets["client_secret"]
    refresh_token = st.secrets["refresh_token"]

    creds = Credentials(
        token=None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=SCOPES,
    )

    creds.refresh(Request())
    return build("drive", "v3", credentials=creds)


def search_file_in_folder(service, filename, folder_id):
    """Tìm file theo tên trong FOLDER_ID cụ thể."""
    query = (
        f"'{folder_id}' in parents and name = '{filename}' and trashed = false"
    )
    results = (
        service.files()
        .list(q=query, spaces="drive", fields="files(id, name)")
        .execute()
    )
    items = results.get("files", [])
    if items:
        return items[0]["id"]
    return None


def extract_file_id_from_link(link):
    """Trích xuất file_id từ URL Google Drive link."""
    if not link:
        return None
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", link)
    if match:
        return match.group(1)
    match_id = re.search(r"id=([a-zA-Z0-9_-]+)", link)
    if match_id:
        return match_id.group(1)
    return None


def delete_file_from_drive(file_link_or_id):
    """Xóa file PDF khỏi Google Drive để dọn dẹp bộ nhớ."""
    try:
        service = get_drive_service()
        if not service or not file_link_or_id:
            return
        
        file_id = extract_file_id_from_link(file_link_or_id) or file_link_or_id
        service.files().delete(fileId=file_id).execute()
    except Exception as e:
        st.warning(f"⚠️ Không thể xóa file trên Drive: {e}")


def upload_pdf_to_drive(uploaded_file):
    try:
        service = get_drive_service()
        if not service:
            st.error("❌ Không thể khởi tạo Google Drive Service.")
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


def save_excel_to_drive(service, wb):
    """Hàm phụ trợ lưu Workbook đè lên Google Drive."""
    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)

    media = MediaIoBaseUpload(
        out_stream,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )

    file_id = search_file_in_folder(service, EXCEL_FILE_NAME, FOLDER_ID)
    if file_id:
        service.files().update(fileId=file_id, media_body=media).execute()
    else:
        file_metadata = {"name": EXCEL_FILE_NAME, "parents": [FOLDER_ID]}
        service.files().create(body=file_metadata, media_body=media).execute()


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

                thuoc_dict[key] = (
                    str(row[1]).strip() if len(row) > 1 and row[1] else ""
                )
                thuoc_quycach_dict[key] = (
                    str(row[2]).strip() if len(row) > 2 and row[2] else ""
                )
                thuoc_chidinh_dict[key] = (
                    str(row[3]).strip() if len(row) > 3 and row[3] else ""
                )
                thuoc_chongchidinh_dict[key] = (
                    str(row[4]).strip() if len(row) > 4 and row[4] else ""
                )

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
                            st.write(f"- **{ten_bd}**: *(Chưa đính kèm PDF)*")
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
# CHỨC NĂNG 3: TRANG THÊM/CẬP NHẬT HOẠT CHẤT & BIỆT DƯỢC
# ==============================================================================
elif chon_tab == "➕ Thêm/Cập Nhật Dữ Liệu Thuốc":
    st.subheader("🔒 Đăng nhập hệ thống quản lý")

    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        password_input = st.text_input(
            "Nhập mật khẩu để truy cập trang này:", type="password"
        )
        if st.button("Đăng nhập"):
            if password_input == "khoaduoc123":
                st.session_state["authenticated"] = True
                st.success("🔓 Đăng nhập thành công!")
                st.rerun()
            else:
                st.error("❌ Mật khẩu không chính xác. Vui lòng thử lại!")
    else:
        st.success("🔑 Đã xác thực thành công dưới quyền Quản lý Khóa dược.")
        if st.button("🔒 Đăng xuất"):
            st.session_state["authenticated"] = False
            st.rerun()

        st.markdown("---")
        st.subheader("➕ Thêm mới / Cập nhật Thuốc & Biệt dược")

        # Quản lý session state
        if "edit_lieu_dung" not in st.session_state:
            st.session_state["edit_lieu_dung"] = ""
        if "edit_quy_cach" not in st.session_state:
            st.session_state["edit_quy_cach"] = ""
        if "edit_chi_dinh" not in st.session_state:
            st.session_state["edit_chi_dinh"] = ""
        if "edit_chong_chi_dinh" not in st.session_state:
            st.session_state["edit_chong_chi_dinh"] = ""
        if "edit_biet_duoc_list" not in st.session_state:
            st.session_state["edit_biet_duoc_list"] = [{"id": 0, "ten": "", "link": ""}]

        CREATE_NEW_OPTION = "[ ➕ Nhập tên Hoạt chất mới... ]"
        options_list = [CREATE_NEW_OPTION] + danh_sach_thuoc

        selected_option = st.selectbox(
            "1. Chọn Hoạt chất đã có từ dữ liệu (hoặc chọn nhập mới): *",
            options=options_list,
            index=0,
            help="Chọn thuốc cũ để chỉnh sửa dữ liệu, hoặc chọn '[ ➕ Nhập tên Hoạt chất mới... ]' để tạo thuốc mới.",
        )

        final_hoat_chat_name = ""

        if selected_option == CREATE_NEW_OPTION:
            final_hoat_chat_name = st.text_input(
                "👉 Nhập Tên Hoạt chất chính thức mới (Cột A) *: ",
                placeholder="Ví dụ: Paracetamol",
                key="new_hoat_chat_input",
            )
            if st.session_state.get("last_selected_option") != CREATE_NEW_OPTION:
                st.session_state["last_selected_option"] = CREATE_NEW_OPTION
                st.session_state["edit_lieu_dung"] = ""
                st.session_state["edit_quy_cach"] = ""
                st.session_state["edit_chi_dinh"] = ""
                st.session_state["edit_chong_chi_dinh"] = ""
                st.session_state["edit_biet_duoc_list"] = [{"id": 0, "ten": "", "link": ""}]
        else:
            final_hoat_chat_name = selected_option
            if st.session_state.get("last_selected_option") != selected_option:
                st.session_state["last_selected_option"] = selected_option
                st.session_state["edit_lieu_dung"] = thuoc_dict.get(selected_option, "")
                st.session_state["edit_quy_cach"] = thuoc_quycach_dict.get(selected_option, "")
                st.session_state["edit_chi_dinh"] = thuoc_chidinh_dict.get(selected_option, "")
                st.session_state["edit_chong_chi_dinh"] = thuoc_chongchidinh_dict.get(selected_option, "")

                bd_old = thuoc_bietduoc_dict.get(selected_option, [])
                new_list = []
                for i, bd in enumerate(bd_old):
                    new_list.append({
                        "id": i,
                        "ten": bd.get("ten", ""),
                        "link": bd.get("link", ""),
                    })
                if not new_list:
                    new_list = [{"id": 0, "ten": "", "link": ""}]
                st.session_state["edit_biet_duoc_list"] = new_list

        if final_hoat_chat_name and selected_option != CREATE_NEW_OPTION:
            st.info(f"🔄 Đang chỉnh sửa dữ liệu cũ của: **{final_hoat_chat_name}**")
        elif final_hoat_chat_name:
            st.success(f"✨ Đang tạo mới dữ liệu cho: **{final_hoat_chat_name}**")

        lieu_dung_input = st.text_input(
            "2. Quy định liều dùng (Cột B - cách nhau bởi dấu ';') *: ",
            key="edit_lieu_dung",
            placeholder="Ví dụ: 10 mg/kg/ngày; 15 mg/kg/ngày",
        )
        quy_cach_input = st.text_input(
            "3. Quy cách hàm lượng (Cột C - cách nhau bởi dấu ';'):",
            key="edit_quy_cach",
            placeholder="Ví dụ: 80; 150; 250; 500",
        )
        chi_dinh_input = st.text_area(
            "4. Chỉ định (Cột D):", key="edit_chi_dinh", height=80
        )
        chong_chi_dinh_input = st.text_area(
            "5. Chống chỉ định (Cột E):", key="edit_chong_chi_dinh", height=80
        )

        st.markdown("---")
        st.write("### 💊 Danh sách Biệt dược & File đính kèm (Cột F)")

        def add_biet_duoc_row():
            max_id = max(
                [item["id"] for item in st.session_state["edit_biet_duoc_list"]],
                default=-1,
            ) + 1
            st.session_state["edit_biet_duoc_list"].append(
                {"id": max_id, "ten": "", "link": ""}
            )

        def remove_and_sync_excel(row_id):
            """Xóa dòng biệt dược, xóa file Drive (nếu có) và cập nhật thẳng vào Excel trên Drive."""
            target_item = None
            for item in st.session_state["edit_biet_duoc_list"]:
                if item["id"] == row_id:
                    target_item = item
                    break

            if not target_item:
                return

            with st.spinner("Đang xóa biệt dược và cập nhật lại file Excel trên Drive..."):
                # 1. Xóa file PDF trên Google Drive nếu có link
                if target_item.get("link"):
                    delete_file_from_drive(target_item["link"])

                # 2. Loại bỏ khỏi danh sách session_state
                st.session_state["edit_biet_duoc_list"] = [
                    item
                    for item in st.session_state["edit_biet_duoc_list"]
                    if item["id"] != row_id
                ]
                if not st.session_state["edit_biet_duoc_list"]:
                    st.session_state["edit_biet_duoc_list"] = [{"id": 0, "ten": "", "link": ""}]

                # 3. Chuẩn bị dữ liệu json biệt dược còn lại
                updated_bd_data = []
                for item in st.session_state["edit_biet_duoc_list"]:
                    t_name = item.get("ten", "").strip()
                    if t_name:
                        updated_bd_data.append({"ten": t_name, "link": item.get("link", "")})

                # 4. Nếu đang chỉnh sửa thuốc cũ (đã có trong Excel), cập nhật trực tiếp Excel
                hoat_chat_name = final_hoat_chat_name.strip()
                if hoat_chat_name and selected_option != CREATE_NEW_OPTION:
                    service = get_drive_service()
                    file_id = search_file_in_folder(service, EXCEL_FILE_NAME, FOLDER_ID)
                    if file_id:
                        request = service.files().get_media(fileId=file_id)
                        file_stream = io.BytesIO()
                        downloader = MediaIoBaseDownload(file_stream, request)
                        done = False
                        while not done:
                            _, done = downloader.next_chunk()
                        file_stream.seek(0)
                        wb = load_workbook(file_stream)
                        ws = wb.active

                        # Tìm dòng và ghi đè lại Cột F
                        for r in range(1, ws.max_row + 1):
                            cell_val = ws.cell(row=r, column=1).value
                            if cell_val and str(cell_val).strip() == hoat_chat_name:
                                ws.cell(row=r, column=6, value=json.dumps(updated_bd_data, ensure_ascii=False))
                                break

                        save_excel_to_drive(service, wb)
                        st.cache_data.clear()

            st.toast("🗑️ Đã xóa biệt dược và đồng bộ Excel trên Drive!")
            st.rerun()

        key_prefix = f"{selected_option}_"

        # Hiển thị danh sách biệt dược
        for idx, item in enumerate(st.session_state["edit_biet_duoc_list"]):
            row_id = item["id"]
            col_bd1, col_bd2, col_bd3 = st.columns([2, 3, 1.5])

            with col_bd1:
                item["ten"] = st.text_input(
                    f"Tên biệt dược #{idx+1}:",
                    value=item.get("ten", ""),
                    key=f"{key_prefix}ten_bd_val_{row_id}",
                    placeholder="Nhập tên biệt dược...",
                )

            with col_bd2:
                uploaded_pdf = st.file_uploader(
                    f"File PDF đính kèm #{idx+1} (Không bắt buộc):",
                    type=["pdf"],
                    key=f"{key_prefix}file_bd_val_{row_id}",
                )
                if item.get("link"):
                    st.markdown(f"📄 [Xem file PDF hiện tại]({item['link']})")

            with col_bd3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️ Xóa & Cập nhật Excel", key=f"{key_prefix}del_bd_{row_id}"):
                    remove_and_sync_excel(row_id)

        st.button("➕ Thêm dòng biệt dược tiếp theo", on_click=add_biet_duoc_row)

        st.markdown("---")

        if st.button("💾 ĐỒNG BỘ VÀ LƯU VÀO GOOGLE DRIVE", type="primary"):
            hoat_chat_name = final_hoat_chat_name.strip()
            lieu_dung_val = lieu_dung_input.strip()

            # --- RÀNG BUỘC 1: BẮT BUỘC TÊN HOẠT CHẤT & QUY ĐỊNH LIỀU DÙNG ---
            if not hoat_chat_name:
                st.error("❌ **Lỗi:** Tên Hoạt chất không được để trống!")
            elif not lieu_dung_val:
                st.error("❌ **Lỗi:** Quy định liều dùng (Cột B) không được để trống!")
            else:
                # --- RÀNG BUỘC 2: NẾU TẢI FILE PDF LÊN THÌ BẮT BUỘC PHẢI CÓ TÊN BIỆT DƯỢC ---
                has_upload_error = False
                for idx, item in enumerate(st.session_state["edit_biet_duoc_list"]):
                    row_id = item["id"]
                    t_bd = st.session_state.get(f"{key_prefix}ten_bd_val_{row_id}", "").strip()
                    f_bd = st.session_state.get(f"{key_prefix}file_bd_val_{row_id}", None)

                    # Chỉ báo lỗi nếu TẢI FILE MỚI LÊN mà KHÔNG NHẬP TÊN
                    if f_bd is not None and not t_bd:
                        st.error(f"❌ **Lỗi ở dòng #{idx+1}:** Bạn đã chọn tải file PDF nhưng chưa nhập Tên biệt dược!")
                        has_upload_error = True

                if not has_upload_error:
                    with st.spinner("Đang lưu dữ liệu và đồng bộ với Google Drive..."):
                        biet_duoc_data = []

                        for item in st.session_state["edit_biet_duoc_list"]:
                            row_id = item["id"]
                            t_bd = st.session_state.get(f"{key_prefix}ten_bd_val_{row_id}", "").strip()
                            f_bd = st.session_state.get(f"{key_prefix}file_bd_val_{row_id}", None)

                            # Cho phép lưu biệt dược ngay cả khi không có file PDF (pdf_link = "")
                            if t_bd:
                                pdf_link = item.get("link", "")
                                if f_bd is not None:
                                    # Xóa file cũ nếu đã tồn tại trước đó
                                    if pdf_link:
                                        delete_file_from_drive(pdf_link)

                                    # Upload file PDF mới
                                    _, pdf_link = upload_pdf_to_drive(f_bd)

                                biet_duoc_data.append({"ten": t_bd, "link": pdf_link or ""})

                        # Đọc workbook hiện tại từ Drive
                        service = get_drive_service()
                        file_id = search_file_in_folder(service, EXCEL_FILE_NAME, FOLDER_ID)
                        
                        if file_id:
                            request = service.files().get_media(fileId=file_id)
                            file_stream = io.BytesIO()
                            downloader = MediaIoBaseDownload(file_stream, request)
                            done = False
                            while not done:
                                _, done = downloader.next_chunk()
                            file_stream.seek(0)
                            wb = load_workbook(file_stream)
                        else:
                            wb = openpyxl.Workbook()

                        ws = wb.active

                        # Cập nhật hoặc thêm mới dòng trong Excel
                        row_found = False
                        for r in range(1, ws.max_row + 1):
                            cell_val = ws.cell(row=r, column=1).value
                            if cell_val and str(cell_val).strip() == hoat_chat_name:
                                ws.cell(row=r, column=2, value=lieu_dung_val)
                                ws.cell(row=r, column=3, value=quy_cach_input.strip())
                                ws.cell(row=r, column=4, value=chi_dinh_input.strip())
                                ws.cell(row=r, column=5, value=chong_chi_dinh_input.strip())
                                ws.cell(row=r, column=6, value=json.dumps(biet_duoc_data, ensure_ascii=False))
                                row_found = True
                                break

                        if not row_found:
                            new_row = ws.max_row + 1
                            ws.cell(row=new_row, column=1, value=hoat_chat_name)
                            ws.cell(row=new_row, column=2, value=lieu_dung_val)
                            ws.cell(row=new_row, column=3, value=quy_cach_input.strip())
                            ws.cell(row=new_row, column=4, value=chi_dinh_input.strip())
                            ws.cell(row=new_row, column=5, value=chong_chi_dinh_input.strip())
                            ws.cell(row=new_row, column=6, value=json.dumps(biet_duoc_data, ensure_ascii=False))

                        # Lưu file Excel mới lên Drive
                        save_excel_to_drive(service, wb)

                        st.cache_data.clear()
                        st.success(f"🎉 Đã lưu thành công dữ liệu thuốc **{hoat_chat_name}** lên Google Drive!")
                        st.rerun()
