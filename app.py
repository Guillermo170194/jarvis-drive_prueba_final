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
# SIDEBAR
# =========================

modulo = st.sidebar.radio(
    "📂 Navegación",
    [
        "🏠 Resumen nacional",
        "🏛 Estado",
        "📚 Documental"
    ]
)

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

# =========================
# AGREGAR FILA GOOGLE SHEETS
# =========================

def guardar_historial_sheets(
    fecha,
    entidad,
    clues,
    tipo,
    archivo,
    link,
    file_id
):

    values = [[
        str(fecha),
        entidad,
        clues,
        tipo,
        archivo,
        link,
        file_id
    ]]
    body = {
        "values": values
    }

    sheets_service.spreadsheets().values().append(
        spreadsheetId=EXCEL_FILE_ID,
        range="HISTORIAL_DOCUMENTAL!A:G",
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()


# =========================
# BORRAR ARCHIVO DRIVE
# =========================

def borrar_archivo_drive(link):

    try:

        file_id = (
            link
            .split("/d/")[1]
            .split("/")[0]
        )

        drive_service.files().update(
            fileId=file_id,
            body={
                "trashed": True
            },
            supportsAllDrives=True
        ).execute()

    except Exception as e:

        st.error(e)
# =========================
# BORRAR FILA SHEETS
# =========================

def borrar_fila_historial(row_number):

    historial = descargar_historial()

    historial = historial.drop(
        historial.index[row_number - 1]
    )

    historial = historial.fillna("")

    # limpiar hoja
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=EXCEL_FILE_ID,
        range="HISTORIAL_DOCUMENTAL!A:G"
    ).execute()

    # volver a escribir
    values = [historial.columns.tolist()] + historial.values.tolist()

    sheets_service.spreadsheets().values().update(
        spreadsheetId=EXCEL_FILE_ID,
        range="HISTORIAL_DOCUMENTAL!A1",
        valueInputOption="USER_ENTERED",
        body={
            "values": values
        }
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
# CARGAR BASES
# =========================

try:

    base_operativa = descargar_base_operativa()

    historial_base = descargar_historial()

except Exception as e:

    st.error(e)

    base_operativa = pd.DataFrame()

    historial_base = pd.DataFrame()

# =========================
# LIMPIAR COLUMNAS
# =========================

base_operativa.columns = (
    base_operativa.columns
    .astype(str)
    .str.strip()
)

# =========================
# LIMPIEZA DATOS
# =========================

base_operativa[
    "CARPETA FÍSCA (Si/no)"
] = (
    base_operativa[
        "CARPETA FÍSCA (Si/no)"
    ]
    .astype(str)
    .str.upper()
    .str.strip()
)

base_operativa[
    "CORRECTO/INCORRECTO"
] = (
    base_operativa[
        "CORRECTO/INCORRECTO"
    ]
    .astype(str)
    .str.upper()
    .str.strip()
)

# =========================
# KPIs
# =========================

correctos = base_operativa[
    (
        base_operativa[
            "CARPETA FÍSCA (Si/no)"
        ] == "SI"
    )
    &
    (
        base_operativa[
            "CORRECTO/INCORRECTO"
        ] == "CORRECTO"
    )
].shape[0]

incorrectos = base_operativa[
    (
        base_operativa[
            "CARPETA FÍSCA (Si/no)"
        ] == "SI"
    )
    &
    (
        base_operativa[
            "CORRECTO/INCORRECTO"
        ] == "INCORRECTO"
    )
].shape[0]

no_entregados = base_operativa[
    base_operativa[
        "CARPETA FÍSCA (Si/no)"
    ] == "NO"
].shape[0]

# =========================
# KPIs VISUALES
# =========================

if modulo == "🏠 Resumen nacional":

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

# =========================
# CATÁLOGOS
# =========================

entidades = sorted(
    base_operativa[
        "ENTIDAD"
    ]
    .dropna()
    .astype(str)
    .unique()
)

# =========================
# DOCUMENTAL
# =========================

if modulo == "📚 Documental":

    # =========================
    # FORMULARIO
    # =========================

    st.markdown("---")

    c1, c2 = st.columns(2)

    with c1:

        entidad = st.selectbox(
            "📍 Entidad",
            entidades
        )

    with c2:

        clues_filtrados = (
            base_operativa[
                base_operativa["ENTIDAD"] == entidad
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
    # GUARDAR DOCUMENTO
    # =========================

    if st.button("📤 Guardar documento"):

        if not archivo:

            st.warning(
                "⚠ Debes subir un archivo"
            )

        else:

            # =========================
            # VALIDAR DUPLICADOS
            # =========================

            historial_actual = descargar_historial()

            existente = historial_actual[
                (
                    historial_actual["CLUES"]
                    .astype(str)
                    == str(clues)
                )
                &
                (
                    historial_actual["Tipo"]
                    .astype(str)
                    == str(tipo)
                )
            ]

            if not existente.empty:

                archivo_existente = (
                    existente.iloc[-1]
                )

                st.warning(
                    f"⚠ Ya existe un documento "
                    f"de tipo '{tipo}' "
                    f"para la CLUES {clues}"
                )

                st.info(
                    f"📄 Archivo actual: "
                    f"{archivo_existente['Archivo']}"
                )

                reemplazar = st.button("♻ Sustituir")

                if not reemplazar:

                    st.stop()

                # =========================
                # MANDAR A PAPELERA
                # =========================

                try:

                    drive_service.files().update(
                        fileId=archivo_existente["File ID"],
                        body={
                            "trashed": True
                        },
                        supportsAllDrives=True
                    ).execute()

                    # =========================
                    # BORRAR REGISTRO HISTORIAL
                    # =========================

                    historial_actual = (
                        historial_actual.drop(
                            existente.index
                        )
                    )

                    historial_actual = (
                        historial_actual.fillna("")
                    )

                    # limpiar hoja
                    (
                        sheets_service.spreadsheets()
                        .values()
                        .clear(
                            spreadsheetId=EXCEL_FILE_ID,
                            range="HISTORIAL_DOCUMENTAL!A:G"
                        )
                        .execute()
                    )

                    # reescribir
                    values = [
                        historial_actual.columns.tolist()
                    ] + historial_actual.values.tolist()

                    (
                        sheets_service.spreadsheets()
                        .values()
                        .update(
                            spreadsheetId=EXCEL_FILE_ID,
                            range="HISTORIAL_DOCUMENTAL!A1",
                            valueInputOption="USER_ENTERED",
                            body={
                                "values": values
                            }
                        )
                        .execute()
                    )

                    st.success(
                        "♻ Documento anterior sustituido"
                    )

                except Exception as e:

                    st.error(e)

                    st.stop()

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

                tipo_archivo = (
                    tipo
                    .replace(" ", "_")
                )

                nombre_drive = (
                    f"{tipo_archivo}_{clues}_{archivo.name}"
                )

                carpeta_entidad = (
                    obtener_carpeta_entidad(
                        entidad
                    )
                )

                file_metadata = {
                    "name": nombre_drive,
                    "parents": [carpeta_entidad]
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

            file_id = uploaded_file["id"]

            drive_link = uploaded_file[
                "webViewLink"
            ]

            os.remove(temp_path)

            guardar_historial_sheets(
                fecha=pd.Timestamp.now(),
                entidad=entidad,
                clues=clues,
                tipo=tipo,
                archivo=archivo.name,
                link=drive_link,
                file_id=file_id
            )

            st.cache_data.clear()

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

        historial = descargar_historial()

        for i, row in historial.iterrows():

            c1, c2, c3, c4, c5, c6 = st.columns(
                [2, 2, 2, 2, 1, 1]
            )

            with c1:
                st.write(row["Fecha"])

            with c2:
                st.write(row["Entidad"])

            with c3:
                st.write(row["CLUES"])

            with c4:
                st.write(row["Tipo"])

            with c5:

                st.link_button(
                    "👁",
                    row["Link"]
                )

            with c6:

                if st.button(
                    "🗑",
                    key=f"delete_{i}"
                ):

                    borrar_archivo_drive(
                        row["Link"]
                    )

                    borrar_fila_historial(
                        i + 1
                    )

                    st.cache_data.clear()

                    st.rerun()

    except Exception as e:

        st.error(e)