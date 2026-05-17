import streamlit as st
import os
import json
import tempfile
import io

import pandas as pd

from google.oauth2 import service_account
from googleapiclient.discovery import build
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

# =========================
# DESCARGAR EXCEL
# =========================

@st.cache_data(ttl=30)
@st.cache_data(ttl=30)
def descargar_base_operativa():

    request = (
        drive_service.files()
        .get_media(
            fileId=EXCEL_FILE_ID
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
        .get_media(
            fileId=EXCEL_FILE_ID
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
# CARGAR HISTORIAL
# =========================

try:

    base_operativa = descargar_base_operativa()
    historial_base = descargar_historial()

    historial_base.columns = (
        historial_base.columns
        .astype(str)
        .str.strip()
    )

except:

    historial_base = pd.DataFrame(
        columns=[
            "Fecha",
            "Entidad",
            "CLUES",
            "Tipo",
            "Archivo",
            "Link"
        ]
    )

# =========================
# LIMPIEZA COLUMNAS
# =========================

base_operativa[
    "CARPETA FÍSCA (Si/no)"
] = (
    historial_base[
        "CARPETA FÍSCA (Si/no)"
    ]
    .astype(str)
    .str.upper()
    .str.strip()
)

historial_base[
    "CORRECTO/INCORRECTO"
] = (
    historial_base[
        "CORRECTO/INCORRECTO"
    ]
    .astype(str)
    .str.upper()
    .str.strip()
)

# =========================
# KPIs PRINCIPALES
# =========================

correctos = historial_base[
    (
        historial_base[
            "CARPETA FÍSCA (Si/no)"
        ] == "SI"
    )
    &
    (
        historial_base[
            "CORRECTO/INCORRECTO"
        ] == "CORRECTO"
    )
].shape[0]

incorrectos = historial_base[
    (
        historial_base[
            "CARPETA FÍSCA (Si/no)"
        ] == "SI"
    )
    &
    (
        historial_base[
            "CORRECTO/INCORRECTO"
        ] == "INCORRECTO"
    )
].shape[0]

no_entregados = historial_base[
    historial_base[
        "CARPETA FÍSCA (Si/no)"
    ] == "NO"
].shape[0]

# =========================
# KPIs VISUALES
# =========================

st.markdown("---")

k1, k2, k3 = st.columns(3)

with k1:

    st.metric(
        "✅ Correctos",
        correctos
    )

with k2:

    st.metric(
        "❌ Incorrectos",
        incorrectos
    )

with k3:

    st.metric(
        "📭 No entregados",
        no_entregados
    )

st.markdown("---")
# =========================
# LIMPIEZA COLUMNAS
# =========================

historial_base[
    "CARPETA FÍSCA (Si/no)"
] = (
    historial_base[
        "CARPETA FÍSCA (Si/no)"
    ]
    .astype(str)
    .str.upper()
    .str.strip()
)

historial_base[
    "CORRECTO/INCORRECTO"
] = (
    historial_base[
        "CORRECTO/INCORRECTO"
    ]
    .astype(str)
    .str.upper()
    .str.strip()
)

# =========================
# KPIs PRINCIPALES
# =========================

correctos = historial_base[
    (
        historial_base[
            "CARPETA FÍSCA (Si/no)"
        ] == "SI"
    )
    &
    (
        historial_base[
            "CORRECTO/INCORRECTO"
        ] == "CORRECTO"
    )
].shape[0]

incorrectos = historial_base[
    (
        historial_base[
            "CARPETA FÍSCA (Si/no)"
        ] == "SI"
    )
    &
    (
        historial_base[
            "CORRECTO/INCORRECTO"
        ] == "INCORRECTO"
    )
].shape[0]

no_entregados = historial_base[
    historial_base[
        "CARPETA FÍSCA (Si/no)"
    ] == "NO"
].shape[0]

# =========================
# KPIs VISUALES
# =========================

st.markdown("---")

k1, k2, k3 = st.columns(3)

with k1:

    st.metric(
        "✅ Correctos",
        correctos
    )

with k2:

    st.metric(
        "❌ Incorrectos",
        incorrectos
    )

with k3:

    st.metric(
        "📭 No entregados",
        no_entregados
    )

st.markdown("---")


# =========================
# CATÁLOGOS
# =========================

entidades = sorted(
    historial_base[
        "Entidad"
    ]
    .dropna()
    .astype(str)
    .unique()
)

# =========================
# FORMULARIO
# =========================

st.markdown("---")

col1, col2 = st.columns(2)

with col1:

    entidad = st.selectbox(
        "📍 Entidad",
        entidades
    )

with col2:

    clues_filtrados = (
        historial_base[
            historial_base["Entidad"] == entidad
        ]["CLUES"]
        .dropna()
        .astype(str)
        .unique()
    )

    clues = st.selectbox(
        "🏥 CLUES",
        sorted(clues_filtrados)
    )

tipo = st.selectbox(
    "📄 Tipo documental",
    [
        "Entrega",
        "Corrección",
        "Primer reiterativo",
        "Segundo reiterativo",
        "Tercer reiterativo",
        "Correo",
        "Otro"
    ]
)

archivo = st.file_uploader(
    "📎 Subir archivo"
)

# =========================
# GUARDAR
# =========================

if st.button("📤 Guardar documento"):

    if not archivo:

        st.warning(
            "⚠ Debes subir un archivo"
        )

    else:

        st.info(
            "Subiendo archivo..."
        )

        with tempfile.NamedTemporaryFile(
            delete=False
        ) as temp_file:

            temp_file.write(
                archivo.getbuffer()
            )

            temp_path = temp_file.name

        nombre_drive = (
            f"{entidad}_{clues}_{archivo.name}"
        )

        file_metadata = {
            "name": nombre_drive,
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

        os.remove(temp_path)

        # =========================
        # NUEVA FILA
        # =========================

        nueva_fila = pd.DataFrame([
            {
                "Fecha": pd.Timestamp.now(),
                "Entidad": entidad,
                "CLUES": clues,
                "Tipo": tipo,
                "Archivo": archivo.name,
                "Link": drive_link
            }
        ])

        historial_guardado = pd.concat(
            [
                historial_base,
                nueva_fila
            ],
            ignore_index=True
        )

        actualizar_excel(
            historial_guardado
        )

        st.success(
            "✅ Documento guardado correctamente"
        )

        st.link_button(
            "📂 Abrir archivo",
            drive_link
        )

# =========================
# HISTORIAL
# =========================

st.markdown("---")

st.markdown(
    "## 📚 Historial documental"
)

try:

    historial = descargar_excel()

    st.dataframe(
        historial,
        use_container_width=True,
        hide_index=True
    )

except:

    st.warning(
        "No se pudo cargar historial"
    )