import json
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import gspread
import pandas as pd
import streamlit as st

# Định nghĩa các quyền truy cập Google Drive & Google Sheets
SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Tên file Google Sheet (hoặc ID) lưu trữ danh sách thuốc
SPREADSHEET_NAME = "Danh_Sach_Thuoc_Lieu"


# ---------------------------------------------------------
# 1. Khởi tạo & Xác thực Google Credentials
# ---------------------------------------------------------
def get_credentials():
    creds = None

    # Lấy từ Streamlit Secrets khi chạy trên Streamlit Cloud
    if "google_drive" in st.secrets and "token" in st.secrets["google_drive"]:
        try:
            token_info = json.loads(st.secrets["google_drive"]["token"])
            creds = Credentials.from_authorized_user_info(token_info, SCOPES)
        except Exception as e:
            st.error(f"⚠️ Lỗi đọc Token từ Secrets: {e}")

    # Lấy từ file local khi chạy dưới máy cá nhân
    elif os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # Tự động refresh token nếu đã hết hạn
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as e:
            st.error(f"❌ Không thể refresh token: {e}")
            creds = None

    # Tự tạo token mới nếu chạy local lần đầu
    if not creds or not creds.valid:
        if os.path.exists("credentials.json"):
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        else:
            st.error(
                "❌ Không tìm thấy thông tin xác thực Google. Vui lòng kiểm tra Streamlit Secrets."
            )
            return None

    return creds


@st.cache_resource
def get_gspread_client():
    creds = get_credentials()
    if creds:
        return gspread.authorize(creds)
    return None


# ---------------------------------------------------------
# 2. Các hàm Đọc & Ghi dữ liệu trên Google Sheets
# ---------------------------------------------------------
def get_or_create_worksheet():
    gc = get_gspread_client()
    if not gc:
        return None

    try:
        # Mở Bảng tính theo tên
        sh = gc.open(SPREADSHEET_NAME)
    except gspread.exceptions.SpreadsheetNotFound:
        # Nếu chưa có thì tự động tạo mới trên Google Drive
        sh = gc.create(SPREADSHEET_NAME)
        worksheet = sh.get_worksheet(0)
        # Tạo hàng tiêu đề mặc định
        worksheet.append_row(["Mã Thuốc", "Tên Thuốc", "Hàm Lượng", "Đơn Vị", "Liều Dùng (mg/kg)", "Ghi Chú"])
        return worksheet

    return sh.get_worksheet(0)


def load_data_from_sheets():
    worksheet = get_or_create_worksheet()
    if not worksheet:
        return pd.DataFrame()

    records = worksheet.get_all_records()
    df = pd.DataFrame(records)
    return df


def save_data_to_sheets(df):
    worksheet = get_or_create_worksheet()
    if worksheet:
        # Xóa dữ liệu cũ và ghi đè DataFrame mới lên Sheet
        worksheet.clear()
        worksheet.update(
            [df.columns.values.tolist()] + df.astype(str).values.tolist()
        )
        st.success("✅ Đã cập nhật thành công lên Google Sheets!")


def add_row_to_sheets(new_row_dict):
    worksheet = get_or_create_worksheet()
    if worksheet:
        worksheet.append_row(list(new_row_dict.values()))
        st.success("✅ Đã thêm thuốc mới vào Google Sheets!")


# ---------------------------------------------------------
# 3. Giao diện ứng dụng Streamlit
# ---------------------------------------------------------
st.set_page_config(
    page_title="Tính Liều Thuốc & Quản Lý Dữ Liệu", layout="wide"
)
st.title("💊 Ứng Dụng Tính Liều Thuốc (Google Sheets API)")

# Load dữ liệu từ Google Sheets
df_thuoc = load_data_from_sheets()

tab1, tab2 = st.tabs(["🧮 Tính Liều Thuốc", "⚙️ Quản Lý Danh Sách Thuốc"])

with tab1:
    st.subheader("Thực Hiện Tính Liều")
    if not df_thuoc.empty and "Tên Thuốc" in df_thuoc.columns:
        selected_drug = st.selectbox(
            "Chọn Thuốc:", df_thuoc["Tên Thuốc"].tolist()
        )
        weight = st.number_input(
            "Cân nặng bệnh nhân (kg):", min_value=0.0, value=10.0, step=0.5
        )

        if st.button("Tính Liều"):
            drug_info = df_thuoc[df_thuoc["Tên Thuốc"] == selected_drug].iloc[0]
            try:
                dose_per_kg = float(drug_info.get("Liều Dùng (mg/kg)", 0))
                total_dose = dose_per_kg * weight
                st.info(
                    f"👉 **Khuyến cáo liều dùng:** {total_dose:.2f} mg (Tính theo liều {dose_per_kg} mg/kg)"
                )
            except ValueError:
                st.error("Dữ liệu liều dùng của thuốc này trong bảng không hợp lệ.")
    else:
        st.warning(
            "Chưa có dữ liệu thuốc. Vui lòng thêm dữ liệu ở tab Quản lý."
        )

with tab2:
    st.subheader("Danh Sách Thuốc Hiện Tại (Từ Google Sheets)")

    if not df_thuoc.empty:
        # Cho phép chỉnh sửa bảng trực tiếp trên giao diện Streamlit
        edited_df = st.data_editor(df_thuoc, num_rows="dynamic", use_container_width=True)

        if st.button("💾 Lưu tất cả thay đổi lên Google Sheets"):
            save_data_to_sheets(edited_df)
            st.cache_data.clear()
    else:
        st.info("Bảng dữ liệu đang trống.")

    st.markdown("---")
    st.subheader("➕ Thêm Nhanh Thuốc Mới")
    with st.form("add_drug_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            ma_thuoc = st.text_input("Mã Thuốc")
            ten_thuoc = st.text_input("Tên Thuốc")
        with col2:
            ham_luong = st.text_input("Hàm Lượng")
            don_vi = st.text_input("Đơn Vị (mg, ml, viên)")
        with col3:
            lieu_dung = st.number_input("Liều Dùng (mg/kg)", min_value=0.0)
            ghi_chu = st.text_input("Ghi Chú")

        submitted = st.form_submit_button("Thêm Thuốc")
        if submitted:
            new_data = {
                "Mã Thuốc": ma_thuoc,
                "Tên Thuốc": ten_thuoc,
                "Hàm Lượng": ham_luong,
                "Đơn Vị": don_vi,
                "Liều Dùng (mg/kg)": lieu_dung,
                "Ghi Chú": ghi_chu,
            }
            add_row_to_sheets(new_data)
            st.rerun()
