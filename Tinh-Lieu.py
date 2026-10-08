import streamlit as st
import pandas as pd
import io
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# Config trang
st.set_page_config(page_title="Quản Lý Dữ Liệu Thuốc", layout="wide")

SCOPES = ['https://www.googleapis.com/auth/drive']

# ----------------------------------------------------
# 1. KẾT NỐI GOOGLE DRIVE API
# ----------------------------------------------------
@st.cache_resource
def get_drive_service():
    try:
        creds = Credentials(
            token=None,
            refresh_token=st.secrets["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=st.secrets["client_id"],
            client_secret=st.secrets["client_secret"],
            scopes=SCOPES
        )
        creds.refresh(Request())
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"Lỗi kết nối Google Drive: {e}")
        return None

service = get_drive_service()

# ----------------------------------------------------
# 2. CÁC HÀM XỬ LÝ FILE TRÊN GOOGLE DRIVE
# ----------------------------------------------------
FILE_EXCEL_NAME = "database_thuoc.xlsx"

def get_file_id_by_name(file_name):
    """Tìm File ID theo tên file trên Google Drive"""
    query = f"name = '{file_name}' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    return None

def load_data_from_excel():
    """Đọc file Excel CSDL từ Google Drive"""
    file_id = get_file_id_by_name(FILE_EXCEL_NAME)
    if not file_id:
        # Nếu chưa có file Excel, trả về DataFrame rỗng với các cột chuẩn
        return pd.DataFrame(columns=[
            "id", "biet_duoc", "hoat_chat", "ham_luong", "don_vi_tinh", 
            "quy_cach", "hang_san_xuat", "nuoc_san_xuat", "pdf_file_id", "pdf_file_name"
        ])
    
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    file_stream.seek(0)
    
    df = pd.read_excel(file_stream)
    return df

def save_data_to_excel(df):
    """Ghi đè DataFrame vào file Excel trên Google Drive"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    file_id = get_file_id_by_name(FILE_EXCEL_NAME)
    media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)

    if file_id:
        service.files().update(fileId=file_id, media_body=media).execute()
    else:
        file_metadata = {'name': FILE_EXCEL_NAME, 'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()

def upload_pdf_to_drive(uploaded_file):
    """Upload file PDF lên Google Drive và trả về File ID"""
    file_metadata = {'name': uploaded_file.
