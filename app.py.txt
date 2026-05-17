import streamlit as st
import os
import json
import tempfile

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# =========================
# CONFIG STREAMLIT
# =========================

st.set_page_config(
    page_title="JARVIS DRIVE",
    layout="wide"
)

# =========================
# ESTILO
# =========================

st.markdown("""
<style>

.stApp{
    background-color:#F4F6F9;
    font-family:'Segoe UI',sans-serif;
}

.titulo{
    font-size:42px;
    font-weight:900;
    color:#9F2241;
}

.subtitulo{
    font-size:18px;
    color:#235B4E;
}

</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================

st.markdown(
    "<div class='titulo'>🧠 JARVIS DRIVE</div>",
    unsafe_allow_html=True
)

st.markdown(
    "<div class='subtitulo'>Carga documental en Google Drive</div>",
    unsafe_allow_html=True
)

# =========================
# GOOGLE DRIVE
# =========================

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

FOLDER_ID = os.environ["FOLDER_ID"]

google_credentials = json.loads(
    os.environ["GOOGLE_CREDENTIALS"]
)

credentials = service_account.Credentials.from_service_account_info(
    google_credentials,
    scopes=SCOPES
)

drive_service = build(
    "drive",
    "v3",
    credentials=credentials
)

# =========================
# SUBIDA
# =========================

archivo = st.file_uploader(
    "📎 Subir archivo"
)

if archivo:

    st.info("Subiendo archivo...")

    with tempfile.NamedTemporaryFile(
        delete=False
    ) as temp_file:

        temp_file.write(
            archivo.getbuffer()
        )

        temp_path = temp_file.name

    file_metadata = {
        "name": archivo.name,
        "parents": [FOLDER_ID]
    }

    media = MediaFileUpload(
        temp_path,
        resumable=True
    )

    uploaded_file = (
        drive_service.files()
        .create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True
        )
        .execute()
    )

    drive_link = uploaded_file[
        "webViewLink"
    ]

    st.success(
        "✅ Archivo subido correctamente"
    )

    st.link_button(
        "📂 Abrir en Drive",
        drive_link
    )

    os.remove(temp_path)