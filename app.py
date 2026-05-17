import streamlit as st
import os
import json
import tempfile
import io

import pandas as pd

from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import (
    MediaFileUpload,
    MediaIoBaseDownload
)

# =========================
# CONFIG STREAMLIT
# =========================

st.set_page_config(
    page_title="JARVIS CLOUD",
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
    "<div class='titulo'>🧠 JARVIS CLOUD</div>",
    unsafe_allow_html=True
)

st.markdown(
    "<div class='subtitulo'>Sistema documental en la nube</div>",
    unsafe_allow_html=True
)

# =========================
# GOOGLE DRIVE
# =========================

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

FOLDER_ID = os.environ[
    "FOLDER_ID"
]

EXCEL_FILE_ID = os.environ[
    "EXCEL_FILE_ID"
]

google_credentials = json.loads(
    os.environ[
        "GOOGLE_CREDENTIALS"
    ]
)

credentials = (
    service_account.Credentials
    .from_service_account_info(
        google_credentials,
        scopes=SCOPES
    )
)

drive_service = build(
    "drive",
    "v3",
    credentials=credentials
)
sheets_service = build(
    "sheets",
    "v4",
    credentials=credentials
)

# =========================
# DESCARGAR EXCEL
# =========================

@st.cache_data(ttl=30)
def descargar_base_operativa():

    request = (
        drive_service.files()
.export_media(
    fileId=EXCEL_FILE_ID,
    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
    )

    archivo = io.BytesIO()

    downloader = MediaIoBaseDownload(
        archivo,
        request
    )

    done = False

    while done is False:

        status, done = downloader.next_chunk()

    archivo.seek(0)

    return pd.read_excel(
        archivo,
        sheet_name=0
    )


@st.cache_data(ttl=30)
def descargar_historial():

    request = (
        drive_service.files()
.export_media(
    fileId=EXCEL_FILE_ID,
    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
    )

    archivo = io.BytesIO()

    downloader = MediaIoBaseDownload(
        archivo,
        request
    )

    done = False

    while done is False:

        status, done = downloader.next_chunk()

    archivo.seek(0)

    return pd.read_excel(
        archivo,
        sheet_name="HISTORIAL_DOCUMENTAL"
    )

# =========================
# ACTUALIZAR EXCEL
# =========================

def actualizar_excel(df):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="HISTORIAL_DOCUMENTAL",
            index=False
        )

    output.seek(0)

    temp_excel = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    )

    temp_excel.write(
        output.read()
    )

    temp_excel.close()

    media = MediaFileUpload(
        temp_excel.name,
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )

    (
        drive_service.files()
        .update(
            fileId=EXCEL_FILE_ID,
            media_body=media
        )
        .execute()
    )

    os.remove(temp_excel.name)
# =========================
# AGREGAR FILA GOOGLE SHEETS
# =========================

def guardar_historial_sheets(
    fecha,
    entidad,
    clues,
    tipo,
    archivo,
    link
):

    values = [[
        str(fecha),
        entidad,
        clues,
        tipo,
        archivo,
        link
    ]]

    body = {
        "values": values
    }

    sheets_service.spreadsheets().values().append(
        spreadsheetId=EXCEL_FILE_ID,
        range="HISTORIAL_DOCUMENTAL!A:F",
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()
# =========================
# BUSCAR O CREAR CARPETA
# =========================

def obtener_carpeta_entidad(entidad):

    query = f"""
    name = '{entidad}'
    and mimeType = 'application/vnd.google-apps.folder'
    and '{FOLDER_ID}' in parents
    and trashed = false
    """

    resultados = (
        drive_service.files()
        .list(
            q=query,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        )
        .execute()
    )

    carpetas = resultados.get(
        "files",
        []
    )

    # =========================
    # SI YA EXISTE
    # =========================

    if carpetas:

        return carpetas[0]["id"]

    # =========================
    # SI NO EXISTE -> CREAR
    # =========================

    metadata = {
        "name": entidad,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [FOLDER_ID]
    }

    carpeta = (
        drive_service.files()
        .create(
            body=metadata,
            fields="id",
            supportsAllDrives=True
        )
        .execute()
    )

    return carpeta["id"]
# =========================
# CARGAR HISTORIAL
# =========================

try:

    historial = descargar_historial()

    for i, row in historial.iterrows():

        c1, c2, c3, c4, c5 = st.columns(
            [2, 2, 2, 2, 1]
        )

        with c1:
            st.write(
                row["Fecha"]
            )

        with c2:
            st.write(
                row["Entidad"]
            )

        with c3:
            st.write(
                row["CLUES"]
            )

        with c4:
            st.write(
                row["Tipo"]
            )

        with c5:

            st.link_button(
                "👁",
                row["Link"]
            )

except:

    st.warning(
        "No se pudo cargar historial"
    )
