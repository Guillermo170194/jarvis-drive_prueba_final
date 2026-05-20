from reportlab.lib.pagesizes import landscape
from reportlab.platypus.flowables import HRFlowable
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)

from reportlab.lib import colors

from reportlab.lib.pagesizes import letter

from reportlab.lib.styles import getSampleStyleSheet

from reportlab.platypus.flowables import Image

from reportlab.pdfbase import pdfmetrics

from reportlab.pdfbase.ttfonts import TTFont

from reportlab.lib.enums import TA_CENTER

from reportlab.lib.styles import ParagraphStyle
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
# SIDEBAR
# =========================

modulo = st.sidebar.radio(
    "📂 Navegación",
    [
        "🏠 Resumen nacional",
        "🏛 Estado",
        "📚 Documental",
        "📦 Inventarios",
        "🕵 Supervisión"
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
# DESCARGAR EXCEL ÚNICO
# =========================

@st.cache_data(ttl=300)
def descargar_excel():

    request = (
        drive_service.files()
        .export_media(
            fileId=EXCEL_FILE_ID,
            mimeType=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )
    )

    archivo = io.BytesIO()

    downloader = MediaIoBaseDownload(
        archivo,
        request
    )

    done = False

    while not done:

        status, done = downloader.next_chunk()

    archivo.seek(0)

    return archivo

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
    fecha_documento,
    link,
    file_id
):

    values = [[
        str(fecha),
        entidad,
        clues,
        tipo,
        archivo,
        str(fecha_documento),
        link,
        file_id
    ]]

    body = {
        "values": values
    }

    sheets_service.spreadsheets().values().append(
        spreadsheetId=EXCEL_FILE_ID,
        range="HISTORIAL_DOCUMENTAL!A:H",
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()

def guardar_inventario_sheets(
    fecha,
    entidad,
    clues,
    tipo,
    archivo,
    inventario_fisico,
    link,
    file_id
):

    values = [[
        str(fecha),
        entidad,
        clues,
        tipo,
        archivo,
        inventario_fisico,
        link,
        file_id
    ]]

    body = {
        "values": values
    }

    sheets_service.spreadsheets().values().append(
        spreadsheetId=EXCEL_FILE_ID,
        range="INVENTARIOS!A:H",
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()


def guardar_supervision_sheets(
    rows
):

    body = {
        "values": rows
    }

    sheets_service.spreadsheets().values().append(

        spreadsheetId=EXCEL_FILE_ID,

        range="SUPERVISION!A:R",

        valueInputOption="USER_ENTERED",

        body=body

    ).execute()# =========================
# HISTORIAL SUPERVISION
# =========================

def guardar_historial_supervision(
    fecha,
    entidad,
    clues,
    almacen,
    verificador,
    pdf_link
):

    values = [[
        str(fecha),
        entidad,
        clues,
        almacen,
        verificador,
        pdf_link
    ]]

    body = {
        "values": values
    }

    sheets_service.spreadsheets().values().append(

        spreadsheetId=EXCEL_FILE_ID,

        range="HISTORIAL_SUPERVISION!A:F",

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

    requests = [

        {
            "deleteDimension": {

                "range": {

                    "sheetId": 856716998,

                    "dimension": "ROWS",

                    "startIndex": row_number,

                    "endIndex": row_number + 1
                }
            }
        }
    ]

    body = {
        "requests": requests
    }

    sheets_service.spreadsheets().batchUpdate(

        spreadsheetId=EXCEL_FILE_ID,

        body=body

    ).execute()

# =========================
# BUSCAR O CREAR CARPETA
# =========================

# =========================
# CARPETA ENTIDAD
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

    if carpetas:

        return carpetas[0]["id"]

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
# CARPETA CLUES
# =========================

def obtener_carpeta_clues(
    entidad,
    clues
):

    carpeta_entidad = (
        obtener_carpeta_entidad(
            entidad
        )
    )

    query = f"""
    name = '{clues}'
    and mimeType = 'application/vnd.google-apps.folder'
    and '{carpeta_entidad}' in parents
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

    if carpetas:

        return carpetas[0]["id"]

    metadata = {
        "name": str(clues),
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [carpeta_entidad]
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
# CARPETA INVENTARIOS
# =========================

def obtener_carpeta_inventarios(
    entidad,
    clues
):

    carpeta_clues = (
        obtener_carpeta_clues(
            entidad,
            clues
        )
    )

    query = f"""
    name = 'INVENTARIOS'
    and mimeType = 'application/vnd.google-apps.folder'
    and '{carpeta_clues}' in parents
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

    if carpetas:

        return carpetas[0]["id"]

    metadata = {
        "name": "INVENTARIOS",
        "mimeType":
            "application/vnd.google-apps.folder",
        "parents": [carpeta_clues]
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
# CARPETA SUPERVISION
# =========================

def obtener_carpeta_supervision(
    entidad,
    clues
):

    carpeta_clues = (
        obtener_carpeta_clues(
            entidad,
            clues
        )
    )

    query = f"""
    name = 'SUPERVISION'
    and mimeType = 'application/vnd.google-apps.folder'
    and '{carpeta_clues}' in parents
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

    if carpetas:

        return carpetas[0]["id"]

    metadata = {
        "name": "SUPERVISION",
        "mimeType":
            "application/vnd.google-apps.folder",
        "parents": [carpeta_clues]
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

    excel_file = descargar_excel()

    base_operativa = pd.read_excel(
        excel_file,
        sheet_name=0
    )

    excel_file.seek(0)

    historial_base = pd.read_excel(
        excel_file,
        sheet_name="HISTORIAL_DOCUMENTAL"
    )

    excel_file.seek(0)

    inventarios_base = pd.read_excel(
        excel_file,
        sheet_name="INVENTARIOS"
    )

    excel_file.seek(0)

    historial_supervision_base = pd.read_excel(
        excel_file,
        sheet_name="HISTORIAL_SUPERVISION"
    )

except Exception as e:

    st.error(e)

    base_operativa = pd.DataFrame()

    historial_base = pd.DataFrame()

    inventarios_base = pd.DataFrame()

    historial_supervision_base = pd.DataFrame()
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
# KPIs NACIONALES
# =========================

inventarios_base = inventarios_base.copy()

# TOTAL CLUES
total_clues = (
    base_operativa["CLUES"]
    .dropna()
    .astype(str)
    .nunique()
)

# CLUES CON ARCHIVO FÍSICO
clues_archivo_fisico = (
    inventarios_base[
        inventarios_base["Tipo"]
        .astype(str)
        == "Evidencia física"
    ]["CLUES"]
    .dropna()
    .astype(str)
    .nunique()
)

# CORRECTOS
correctos = (
    base_operativa[
        base_operativa[
            "CORRECTO/INCORRECTO"
        ]
        .astype(str)
        .str.upper()
        == "CORRECTO"
    ]["CLUES"]
    .dropna()
    .astype(str)
    .nunique()
)

# NO ENTREGADOS
no_entregados = (
    total_clues
    - clues_archivo_fisico
)
# =========================
# LISTADOS OPERATIVOS
# =========================

df_correctos = base_operativa[
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
]

df_no_entregados = base_operativa[
    base_operativa[
        "CARPETA FÍSCA (Si/no)"
    ] == "NO"
]

# =========================
# KPIs DOCUMENTALES
# =========================

historial_docs = historial_base.copy()

total_documentos = historial_docs.shape[0]

clues_con_documento = (
    historial_docs["CLUES"]
    .astype(str)
    .nunique()
)

total_clues = (
    base_operativa["CLUES"]
    .astype(str)
    .nunique()
)

clues_pendientes = (
    total_clues - clues_con_documento
)

entregas = historial_docs[
    historial_docs["Tipo"] == "Entrega"
].shape[0]

reiterativos = historial_docs[
    historial_docs["Tipo"]
    .astype(str)
    .str.contains(
        "reiterativo",
        case=False,
        na=False
    )
].shape[0]

	
# =========================
# RESUMEN DOCUMENTAL
# =========================

# =========================
# INVENTARIO FÍSICO RECIBIDO
# =========================

inventarios_base = inventarios_base.copy()

clues_con_inventario = (
    inventarios_base[
        inventarios_base["Tipo"]
        .astype(str)
        == "Evidencia física"
    ]["CLUES"]
    .dropna()
    .astype(str)
    .unique()
)
# =========================
# MATRIZ DOCUMENTAL NACIONAL
# =========================

matriz_documental = []

tipos_documentales = [
    "Entrega",
    "Entrega UAS Y OIC",
    "Corrección",
    "Prórroga",
    "Primer reiterativo",
    "Segundo reiterativo",
    "Tercer reiterativo",
    "Cuarto reiterativo"
]

for _, row_base in base_operativa.iterrows():

    entidad = str(
        row_base["ENTIDAD"]
    )

    clues = str(
        row_base["CLUES"]
    )

    almacen = str(
        row_base["ALMACÉN"]
    )

    historial_clues = historial_docs[
        historial_docs["CLUES"]
        .astype(str)
        == clues
    ]
    supervision_clues = historial_supervision_base[
        historial_supervision_base["CLUES"]
        .astype(str)
        == clues
    ]

    fila = {

        "Entidad": entidad,

        "CLUES": clues,

        "ALMACÉN": almacen,

        "INVENTARIO FÍSICO": (
            "SI"
            if clues in clues_con_inventario
            else "NO"
        ),

        "SUPERVISIÓN": (
            "SI"
            if not supervision_clues.empty
            else "NO"
        )
    }

    for tipo_doc in tipos_documentales:

        documento = historial_clues[
            historial_clues["Tipo"]
            .astype(str)
            == tipo_doc
        ]

        if documento.empty:

            fila[tipo_doc] = "-"
            fila[f"Fecha {tipo_doc}"] = "-"

        else:

            ultimo = documento.iloc[-1]

            fila[tipo_doc] = (
                ultimo["Archivo"]
            )

            fecha_doc = ultimo.iloc[5]

            if (
                pd.isna(fecha_doc)
                or
                str(fecha_doc).strip() == ""
            ):

                fila[f"Fecha {tipo_doc}"] = "-"

            else:

                fila[f"Fecha {tipo_doc}"] = (
                    pd.to_datetime(
                        fecha_doc
                    ).strftime("%d/%m/%Y")
                )

    matriz_documental.append(
        fila
    )
df_resumen_entidad = pd.DataFrame(
    matriz_documental
)
# =========================
# KPIs VISUALES
# =========================

if modulo == "🏠 Resumen nacional":

    st.markdown("---")

    # =========================
    # KPIs PRINCIPALES
    # =========================

    k1, k2, k3, k4 = st.columns(4)

    with k1:

        st.metric(
            "🏥 Total CLUES",
            total_clues
        )

    with k2:

        st.metric(
            "📦 Archivo físico",
            clues_archivo_fisico
        )

    with k3:

        st.metric(
            "✅ Correctos",
            correctos
        )

    with k4:

        st.metric(
            "📭 No entregados",
            no_entregados
        )

    st.markdown("---")

    # =========================
    # KPIs DOCUMENTALES
    # =========================

    d1, d2, d3, d4, d5 = st.columns(5)

    with d1:

        st.metric(
            "📄 Documentos",
            total_documentos
        )

    with d2:

        st.metric(
            "🏥 CLUES con docs",
            clues_con_documento
        )

    with d3:

        st.metric(
            "📭 Pendientes",
            clues_pendientes
        )

    with d4:

        st.metric(
            "✅ Entregas",
            entregas
        )

    with d5:

        st.metric(
            "♻ Reiterativos",
            reiterativos
        )

    st.markdown("---")

    # =========================
    # MATRIZ DOCUMENTAL
    # =========================

    st.markdown(
        "## 📋 Matriz documental nacional"
    )

    st.dataframe(
        df_resumen_entidad,
        use_container_width=True,
        hide_index=True
    )

    st.markdown("---")

    # =========================
    # TABLAS OPERATIVAS
    # =========================

    c1, c2 = st.columns(2)

    with c1:

        st.markdown(
            "## ✅ Correctos"
        )

        st.dataframe(
            df_correctos[
                [
                    "ENTIDAD",
                    "CLUES",
                    "ALMACÉN"
                ]
            ],
            use_container_width=True,
            hide_index=True,
            height=300
        )

    with c2:

        st.markdown(
            "## 📭 No entregados"
        )

        st.dataframe(
            df_no_entregados[
                [
                    "ENTIDAD",
                    "CLUES",
                    "ALMACÉN"
                ]
            ],
            use_container_width=True,
            hide_index=True,
            height=300
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

# =========================
# ESTADO
# =========================

if modulo == "🏛 Estado":

    st.markdown("---")

    entidad_estado = st.selectbox(
        "🏛 Selecciona entidad",
        entidades
    )
    base_estado = base_operativa[
        base_operativa["ENTIDAD"]
        .astype(str)
        == entidad_estado
    ]

    historial_estado = historial_docs[
        historial_docs["Entidad"]
        .astype(str)
        == entidad_estado
    ]
    total_clues_estado = (
        base_estado["CLUES"]
        .astype(str)
        .nunique()
    )

    entregas_estado = historial_estado[
        historial_estado["Tipo"]
        == "Entrega"
    ].shape[0]

    correcciones_estado = historial_estado[
        historial_estado["Tipo"]
        == "Corrección"
    ].shape[0]

    reiterativos_estado = historial_estado[
        historial_estado["Tipo"]
        .astype(str)
        .str.contains(
            "reiterativo",
            case=False,
            na=False
        )
    ].shape[0]
    st.markdown("---")

    e1, e2, e3, e4 = st.columns(4)

    with e1:

        st.metric(
            "🏥 CLUES",
            total_clues_estado
        )

    with e2:

        st.metric(
            "✅ Entregas",
            entregas_estado
        )

    with e3:

        st.metric(
            "🟡 Correcciones",
            correcciones_estado
        )

    with e4:

        st.metric(
            "🔴 Reiterativos",
            reiterativos_estado
        )
    st.markdown("---")

    st.markdown(
        f"## 📋 CLUES - {entidad_estado}"
    )
    resumen_docs = (
        historial_estado
        .groupby("CLUES")["Tipo"]
        .apply(
            lambda x: ", ".join(
                sorted(
                    x.astype(str).unique()
                )
            )
        )
        .reset_index()
    )

    tabla_estado = base_estado[
        [
            "CLUES",
            "ALMACÉN",
            "CORRECTO/INCORRECTO"
        ]
    ]

    tabla_estado = tabla_estado.merge(
        resumen_docs,
        on="CLUES",
        how="left"
    )

    tabla_estado = tabla_estado.rename(
        columns={
            "Tipo": "DOCUMENTOS"
        }
    )

    tabla_estado = tabla_estado.fillna(
        "Sin documentos"
    )
    st.dataframe(
        tabla_estado,
        use_container_width=True,
        hide_index=True,
        height=500
    )
    st.markdown("---")

    st.markdown(
        f"## 📚 Documentos - {entidad_estado}"
    )
    for i, row in historial_estado.iterrows():

        c1, c2, c3, c4 = st.columns(
            [2, 2, 2, 1]
        )

        with c1:

            st.write(
                row["CLUES"]
            )

        with c2:

            st.write(
                row["Tipo"]
            )

        with c3:

            st.write(
                row["Archivo"]
            )

        with c4:

            st.link_button(
                "👁 Abrir",
                row["Link"]
            )

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
    resultado_clues = base_operativa[
        base_operativa["CLUES"]
        .astype(str)
        == str(clues)
    ]

    if resultado_clues.empty:

        almacen_clues = "SIN ALMACÉN"

    else:

        almacen_clues = (
            resultado_clues[
                "ALMACÉN"
            ]
            .astype(str)
            .iloc[0]
        )

    st.info(
        f"🏬 ALMACÉN: {almacen_clues}"
    )

    historial_clues = historial_docs[
        historial_docs["CLUES"]
        .astype(str)
        == str(clues)
    ]

    if historial_clues.empty:

        st.warning(
            "📭 Sin documentos registrados"
        )

    else:

        tipos_documentales = (
            historial_clues["Tipo"]
            .astype(str)
            .unique()
        )

        resumen_documental = ", ".join(
            tipos_documentales
        )

        st.success(
            f"📋 Documentos registrados: "
            f"{resumen_documental}"
        )

    tipo = st.selectbox(
        "📄 Tipo documental",
        [
            "Entrega",
            "Corrección",
            "Prórroga",
            "Entrega UAS Y OIC",
            "Primer reiterativo",
            "Segundo reiterativo",
            "Tercer reiterativo",
            "Cuarto reiterativo",
            "Correo",
            "Otro"
        ]
    )

    archivo = st.file_uploader(
        "📎 Subir archivo"
    )
    fecha_oficio = st.date_input(
        "📅 Fecha del oficio / recepción"
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

            historial_actual = historial_base.copy()

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

                    indice = (
                        existente.index[0]
                    )

                    borrar_fila_historial(
                        indice + 1
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
                    obtener_carpeta_clues(
                        entidad,
                        clues
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
                fecha_documento=fecha_oficio,
                link=drive_link,
                file_id=file_id
            )
            st.cache_data.clear()

            st.toast(
                "✅ Documento guardado correctamente",
                icon="✅"
            )

            st.rerun()

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

        historial = historial_base.copy()

        if historial.empty:

            st.warning(
                "Sin historial documental"
            )

        else:

            historial_tabla = historial[
                [
                    "Fecha",
                    "Entidad",
                    "CLUES",
                    "Tipo",
                    "Archivo"
                ]
            ].copy()

            st.dataframe(
                historial_tabla,
                use_container_width=True,
                hide_index=True,
                height=500
            )

            st.markdown("---")

            st.markdown(
                "### 🗑 Eliminar documento"
            )

            opciones_eliminar = (
                historial[
                    "Archivo"
                ]
                .astype(str)
                .tolist()
            )

            archivo_eliminar = st.selectbox(
                "Selecciona archivo",
                opciones_eliminar
            )

            if st.button(
                "🗑 Eliminar seleccionado"
            ):

                fila_eliminar = historial[
                    historial["Archivo"]
                    .astype(str)
                    == str(archivo_eliminar)
                ].iloc[0]

                borrar_archivo_drive(
                    fila_eliminar["Link"]
                )

                indice = (
                    historial[
                        historial["Archivo"]
                        .astype(str)
                        == str(archivo_eliminar)
                    ].index[0]
                )

                borrar_fila_historial(
                    indice + 1
                )

                st.cache_data.clear()

                st.rerun()

    except Exception as e:

        st.error(e)

# =========================
# DESCARGAR LOGO
# =========================

def descargar_logo():

    file_id = os.environ[
        "LOGO_IMSS_FILE_ID"
    ]

    request = (
        drive_service.files()
        .get_media(
            fileId=file_id,
            supportsAllDrives=True
        )
    )

    temp_logo = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".png"
    )

    downloader = MediaIoBaseDownload(
        temp_logo,
        request
    )

    done = False

    while done is False:

        status, done = downloader.next_chunk()

    temp_logo.close()

    return temp_logo.name

# =========================
# GENERAR PDF SUPERVISIÓN
# =========================

def generar_pdf_supervision(
    entidad,
    clues,
    almacen,
    fecha_supervision,
    nombre_verificador,
    cargo_verificador,
    nombre_almacen,
    cargo_almacen,
    conceptos_generales,
    conceptos_diferencias,
    firma_verificador=None,
    firma_almacen=None
):

    nombre_pdf = (
        f"SUPERVISION_{clues}_{fecha_supervision}.pdf"
    )

    doc = SimpleDocTemplate(
        nombre_pdf,
        pagesize=landscape(letter),
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    elementos = []

    # =========================
    # ENCABEZADO INSTITUCIONAL
    # =========================

    try:

        logo_path = descargar_logo()

        encabezado = Image(
            logo_path,
            width=720,
            height=45
        )

        elementos.append(
            encabezado
        )

    except Exception as e:

        st.error(e)

    elementos.append(
        Spacer(1, 10)
    )

    styles = getSampleStyleSheet()
    estilo_obs = ParagraphStyle(
        "observaciones",
        fontSize=7,
        leading=9
    )

    # =========================
    # TÍTULO
    # =========================

    titulo = Paragraph(
        """
        <font size=14>
        <b>
        SERVICIOS DE SALUD DEL INSTITUTO MEXICANO
        DEL SEGURO SOCIAL PARA EL BIENESTAR
        </b>
        </font>
        <br/>
        <font size=12>
        CÉDULA DE SUPERVISIÓN INVENTARIOS ANUALES
        </font>
        """,
        ParagraphStyle(
            "titulo",
            alignment=TA_CENTER,
            leading=18
        )
    )

    elementos.append(
        titulo
    )

    elementos.append(
        Spacer(1, 8)
    )

    elementos.append(
        HRFlowable(
            width="100%",
            thickness=1.5,
            color=colors.black
        )
    )

    elementos.append(
        Spacer(1, 12)
    )

    # =========================
    # DATOS GENERALES
    # =========================

    datos = [
        [
            "Entidad",
            entidad,
            "CLUES",
            clues
        ],
        [
            "Unidad",
            almacen,
            "Fecha",
            str(fecha_supervision)
        ]
    ]

    tabla_datos = Table(
        datos,
        colWidths=[80, 260, 80, 180]
    )

    tabla_datos.setStyle(
        TableStyle([

            (
                "BACKGROUND",
                (0,0),
                (-1,0),
                colors.lightgrey
            ),

            (
                "GRID",
                (0,0),
                (-1,-1),
                1,
                colors.black
            ),

            (
                "FONTNAME",
                (0,0),
                (-1,-1),
                "Helvetica-Bold"
            ),

            (
                "FONTSIZE",
                (0,0),
                (-1,-1),
                9
            ),

            (
                "VALIGN",
                (0,0),
                (-1,-1),
                "MIDDLE"
            )
        ])
    )

    elementos.append(
        tabla_datos
    )

    elementos.append(
        Spacer(1, 15)
    )

    # =========================
    # SERVIDORES PÚBLICOS
    # =========================

    verificadores = [

        [
            "Servidor público verificador",
            nombre_verificador,
            cargo_verificador
        ],

        [
            "Servidor público almacén",
            nombre_almacen,
            cargo_almacen
        ]
    ]

    tabla_verificador = Table(
        verificadores,
        colWidths=[240, 240, 220]
    )

    tabla_verificador.setStyle(
        TableStyle([

            (
                "GRID",
                (0,0),
                (-1,-1),
                1,
                colors.black
            ),

            (
                "BACKGROUND",
                (0,0),
                (0,-1),
                colors.lightgrey
            ),

            (
                "FONTNAME",
                (0,0),
                (-1,-1),
                "Helvetica"
            ),

            (
                "FONTSIZE",
                (0,0),
                (-1,-1),
                9
            )
        ])
    )

    elementos.append(
        tabla_verificador
    )

    elementos.append(
        Spacer(1, 18)
    )

    # =========================
    # DIAGNÓSTICO GENERAL
    # =========================

    encabezado_general = [[
        "Concepto",
        "Contiene",
        "Piezas",
        "Monto",
        "Firmado",
        "Observaciones"
    ]]

    filas_general = []

    for concepto in conceptos_generales:

        filas_general.append([

            concepto,

            st.session_state[
                f"{concepto}_contiene"
            ],

            "{:,.0f}".format(
                st.session_state[
                    f"{concepto}_piezas"
                ]
            ),

            "${:,.2f}".format(
                st.session_state[
                    f"{concepto}_monto"
                ]
            ),

            st.session_state[
                f"{concepto}_firmado"
            ],

            Paragraph(
                str(
                    st.session_state[
                        f"{concepto}_obs"
                    ]
                ),
                estilo_obs
            )
        ])

    tabla_general = Table(
        encabezado_general + filas_general,
        colWidths=[
            220,
            70,
            90,
            110,
            70,
            180
        ]
    )

    tabla_general.setStyle(
        TableStyle([

            (
                "BACKGROUND",
                (0,0),
                (-1,0),
                colors.black
            ),

            (
                "TEXTCOLOR",
                (0,0),
                (-1,0),
                colors.white
            ),

            (
                "GRID",
                (0,0),
                (-1,-1),
                1,
                colors.black
            ),

            (
                "FONTNAME",
                (0,0),
                (-1,0),
                "Helvetica-Bold"
            ),

            (
                "FONTSIZE",
                (0,0),
                (-1,-1),
                8
            ),

            (
                "VALIGN",
                (0,0),
                (-1,-1),
                "MIDDLE"
            ),

            (
                "WORDWRAP",
                (0,0),
                (-1,-1),
                True
            )

        ])
    )

    elementos.append(
        tabla_general
    )
    elementos.append(
        Spacer(1, 20)
    )

    # =========================
    # DIFERENCIAS
    # =========================

    encabezado_dif = [[
        "Concepto",
        "Existe",
        "Dif más",
        "Dif menos",
        "Observaciones"
    ]]

    filas_dif = []

    for concepto in conceptos_diferencias:

        filas_dif.append([

            concepto,

            st.session_state[
                f"validacion_{concepto}_existe"
            ],

            "{:,.2f}".format(
                st.session_state[
                    f"{concepto}_mas"
                ]
            ),

            "{:,.2f}".format(
                st.session_state[
                    f"{concepto}_menos"
                ]
            ),

            Paragraph(
                str(
                    st.session_state[
                        f"validacion_{concepto}_obs"
                    ]
                ),
                estilo_obs
            )
        ])

    tabla_dif = Table(
        encabezado_dif + filas_dif,
        colWidths=[
            320,
            80,
            100,
            100,
            170
        ]
    )

    tabla_dif.setStyle(
        TableStyle([

            (
                "BACKGROUND",
                (0,0),
                (-1,0),
                colors.black
            ),

            (
                "TEXTCOLOR",
                (0,0),
                (-1,0),
                colors.white
            ),

            (
                "GRID",
                (0,0),
                (-1,-1),
                1,
                colors.black
            ),

            (
                "FONTNAME",
                (0,0),
                (-1,0),
                "Helvetica-Bold"
            ),

            (
                "FONTSIZE",
                (0,0),
                (-1,-1),
                8
            ),

            (
                "WORDWRAP",
                (0,0),
                (-1,-1),
                True
            )

        ])
    )

    elementos.append(
        tabla_dif
    )
    elementos.append(
        Spacer(1, 10)
    )

    # =========================
    # FIRMAS
    # =========================

    firma_img_1 = ""
    firma_img_2 = ""

    try:

        if firma_verificador is not None:

            firma_temp_1 = tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".png"
            )

            firma_temp_1.write(
                firma_verificador.getbuffer()
            )

            firma_temp_1.close()

            firma_img_1 = Image(
                firma_temp_1.name,
                width=140,
                height=50
            )

    except:

        firma_img_1 = ""

    try:

        if firma_almacen is not None:

            firma_temp_2 = tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".png"
            )

            firma_temp_2.write(
                firma_almacen.getbuffer()
            )

            firma_temp_2.close()

            firma_img_2 = Image(
                firma_temp_2.name,
                width=140,
                height=50
            )

    except:

        firma_img_2 = ""

    firmas = Table(

        [
            [
                firma_img_1,
                firma_img_2
            ],

            [
                "___________________________",
                "___________________________"
            ],

            [
                nombre_verificador,
                nombre_almacen
            ]
        ],

        colWidths=[300, 300]
    )

    firmas.setStyle(
        TableStyle([

            (
                "ALIGN",
                (0,0),
                (-1,-1),
                "CENTER"
            ),

            (
                "FONTSIZE",
                (0,0),
                (-1,-1),
                9
            ),

            (
                "TOPPADDING",
                (0,0),
                (-1,-1),
                2
            )
        ])
    )

    elementos.append(
        firmas
    )

    doc.build(
        elementos
    )

    return nombre_pdf

# =========================
# SUPERVISIÓN
# =========================

if modulo == "🕵 Supervisión":

    st.markdown("---")

    st.markdown(
        "# 🕵 Cédula de supervisión"
    )

    s1, s2 = st.columns(2)

    with s1:

        entidad_sup = st.selectbox(
            "📍 Entidad",
            entidades,
            key="sup_entidad"
        )

    with s2:

        clues_disponibles = sorted(
            base_operativa[
                base_operativa["ENTIDAD"]
                .astype(str)
                == str(entidad_sup)
            ]["CLUES"]
            .dropna()
            .astype(str)
            .unique()
        )

        clues_sup = st.selectbox(
            "🏥 CLUES",
            clues_disponibles
        )

    resultado_sup = base_operativa[
        base_operativa["CLUES"]
        .astype(str)
        == str(clues_sup)
    ]

    if resultado_sup.empty:

        almacen_sup = "SIN ALMACÉN"

    else:

        almacen_sup = (
            resultado_sup[
                "ALMACÉN"
            ]
            .astype(str)
            .iloc[0]
        )

    st.info(
        f"🏬 ALMACÉN: {almacen_sup}"
    )
    with st.form("form_supervision"):

        fecha_supervision = st.date_input(
            "📅 Fecha supervisión"
        )

        st.markdown("---")

        # =========================
        # SERVIDORES PÚBLICOS
        # =========================

        st.markdown(
            "## 👤 Servidor público verificador"
        )

        v1, v2 = st.columns(2)

        with v1:

            nombre_verificador = st.text_input(
                "Nombre verificador",
                value="Guillermo Ortega Carteño"
            )

        with v2:

            cargo_verificador = st.text_input(
                "Cargo verificador",
                value="Supervisor de Procesos"
            )

        st.markdown("---")

        st.markdown(
            "## 🏬 Servidor público del almacén"
        )

        a1, a2 = st.columns(2)

        with a1:

            nombre_almacen = st.text_input(
                "Nombre responsable almacén"
            )

        with a2:

            cargo_almacen = st.text_input(
                "Cargo responsable almacén"
            )

        st.markdown("---")

        # =========================
        # DIAGNÓSTICO GENERAL
        # =========================

        st.markdown(
            "## 📋 Diagnóstico general"
        )

        conceptos_generales = [

            "ANEXO 1.- LISTADO GRAL",

            "ANEXO 2.- IMSS B",

            "ANEXO 3.- PROPIOS",

            "Acta de conclusión",

            "ANEXO 5.- ACTA DE INICIO",

            "ANEXO 6.- REPORTE LENTO Y NULO",

            "ANEXO 7.- REPORTE CADUCOS",

            "ANEXO 8.- REP DE PROX A CADUCAR"
        ]

        for concepto in conceptos_generales:

            st.markdown(
                f"### {concepto}"
            )

            c1, c2, c3, c4 = st.columns(4)

            with c1:

                st.selectbox(
                    "Contiene",
                    ["SI", "NO"],
                    key=f"{concepto}_contiene"
                )

            with c2:

                st.number_input(
                    "Piezas",
                    min_value=0.0,
                    step=1.0,
                    format="%.0f",
                    key=f"{concepto}_piezas"
                )

            with c3:

                st.number_input(
                    "Monto",
                    min_value=0.0,
                    step=1.0,
                    format="%.2f",
                    key=f"{concepto}_monto"
                )

            with c4:

                st.selectbox(
                    "Firmado",
                    ["SI", "NO"],
                    key=f"{concepto}_firmado"
                )

            st.text_area(
                "Observaciones",
                key=f"{concepto}_obs"
            )

            st.markdown("---")

        # =========================
        # DIFERENCIAS
        # =========================

        st.markdown(
            "## ⚠ Reporte diferencias"
        )

        conceptos_diferencias = [

            "ANEXO 4.- REPORTE DE DIF"
        ]

        for concepto in conceptos_diferencias:

            st.markdown(
                f"### {concepto}"
            )

            d1, d2, d3 = st.columns(3)

            with d1:

                st.selectbox(
                    "Existe",
                    ["SI", "NO"],
                    key=f"validacion_{concepto}_existe"
                )

            with d2:

                st.number_input(
                    "Dif más",
                    min_value=0.0,
                    step=1.0,
                    format="%.2f",
                    key=f"{concepto}_mas"
                )

            with d3:

                st.number_input(
                    "Dif menos",
                    min_value=0.0,
                    step=1.0,
                    format="%.2f",
                    key=f"{concepto}_menos"
                )

            st.text_area(
                "Observaciones",
                key=f"validacion_{concepto}_obs"
            )

            st.markdown("---")

        # =========================
        # VALIDACIÓN DOCUMENTAL
        # =========================

        st.markdown(
            "## 📄 Validación documental"
        )

        conceptos_validacion = [

            "EXCEL",

            "PDF",

            "FISICO"
        ]

        for concepto in conceptos_validacion:

            st.markdown(
                f"### {concepto}"
            )

            st.selectbox(
                "Existe",
                ["SI", "NO"],
                key=f"{concepto}_existe"
            )

            st.text_area(
                "Observaciones",
                key=f"{concepto}_obs_2"
            )

            st.markdown("---")

        # =========================
        # FIRMAS
        # =========================

        st.markdown(
            "## ✍ Firmas"
        )

        firma_verificador = st.file_uploader(
            "Firma verificador",
            type=["png", "jpg", "jpeg"],
            key="firma_verificador"
        )

        firma_almacen = st.file_uploader(
            "Firma almacén",
            type=["png", "jpg", "jpeg"],
            key="firma_almacen"
        )

        st.markdown("---")

        guardar_supervision = st.form_submit_button(
            "📤 Guardar supervisión"
        )

    # =========================
    # GUARDAR SUPERVISIÓN
    # =========================

    if guardar_supervision:

        rows = []

        # =========================
        # ANEXOS GENERALES
        # =========================

        for concepto in conceptos_generales:

            rows.append([

                str(fecha_supervision),

                entidad_sup,

                clues_sup,

                almacen_sup,

                str(nombre_verificador),

                str(cargo_verificador),

                str(nombre_almacen),

                str(cargo_almacen),

                concepto,

                st.session_state[
                    f"{concepto}_contiene"
                ],

                st.session_state[
                    f"{concepto}_piezas"
                ],

                st.session_state[
                    f"{concepto}_monto"
                ],

                st.session_state[
                    f"{concepto}_firmado"
                ],

                "",

                "",

                "",

                st.session_state[
                    f"{concepto}_obs"
                ]
            ])

        for concepto in conceptos_diferencias:

            rows.append([

                str(fecha_supervision),

                entidad_sup,

                clues_sup,

                almacen_sup,

                str(nombre_verificador),

                str(cargo_verificador),

                str(nombre_almacen),

                str(cargo_almacen),

                concepto,

                "",

                "",

                "",

                "",

                st.session_state[
                     f"{concepto}_contiene"
                ],

                st.session_state[
                    f"{concepto}_mas"
                ],

                st.session_state[
                    f"{concepto}_menos"
                ],

                st.session_state[
                    f"validacion_{concepto}_obs"
                ]
            ])

        for concepto in conceptos_validacion:

            rows.append([

                str(fecha_supervision),

                entidad_sup,

                clues_sup,

                almacen_sup,

                str(nombre_verificador),

                str(cargo_verificador),

                str(nombre_almacen),

                str(cargo_almacen),

                concepto,

                "",

                "",

                "",

                "",

                st.session_state[
                    f"{concepto}_existe"
                ],

                0,

                0,

                st.session_state[
                    f"{concepto}_obs_2"
                ]
            ])

        guardar_supervision_sheets(
            rows
        )
        pdf_generado = generar_pdf_supervision(

            entidad=entidad_sup,

            clues=clues_sup,

            almacen=almacen_sup,

            fecha_supervision=fecha_supervision,

            nombre_verificador=nombre_verificador,

            cargo_verificador=cargo_verificador,

            nombre_almacen=nombre_almacen,

            cargo_almacen=cargo_almacen,

            conceptos_generales=conceptos_generales,

            conceptos_diferencias=conceptos_diferencias,

            firma_verificador=firma_verificador,

            firma_almacen=firma_almacen
        )
        folder_id = obtener_carpeta_supervision(
            entidad_sup,
            clues_sup
        )

        file_metadata = {
            "name": pdf_generado,
            "parents": [folder_id]
        }

        media = MediaFileUpload(
            pdf_generado,
            mimetype="application/pdf"
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

        pdf_link = uploaded_file[
            "webViewLink"
        ]
        if os.path.exists(pdf_generado):

            os.remove(pdf_generado)

        guardar_historial_supervision(

            fecha=fecha_supervision,

            entidad=entidad_sup,

            clues=clues_sup,

            almacen=almacen_sup,

            verificador=nombre_verificador,

            pdf_link=pdf_link
        )

        st.cache_data.clear()

        st.toast(
            "✅ Supervisión guardada correctamente",
            icon="✅"
        )

        st.rerun()
    st.markdown("---")

    st.markdown(
        "## 🕵 Historial supervisiones"
    )

    try:

        historial_supervision = (
            historial_supervision_base.copy()
        )

        if historial_supervision.empty:

            st.warning(
                "Sin supervisiones"
            )

        else:

            historial_supervision = (
                historial_supervision.sort_values(
                    "Fecha",
                    ascending=False
                )
            )

            buscador = st.text_input(
                "🔎 Buscar supervisión"
            )

            if buscador:

                historial_supervision = (
                    historial_supervision[
                        historial_supervision
                        .astype(str)
                        .apply(
                            lambda x:
                            x.str.contains(
                                buscador,
                                case=False,
                                na=False
                            )
                        )
                        .any(axis=1)
                    ]
                )

            for i, row in historial_supervision.iterrows():

                c1, c2, c3, c4, c5 = st.columns(
                    [2,2,3,2,1]
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
                        row["Almacen"]
                    )

                with c4:

                    st.write(
                        row["Verificador"]
                    )

                with c5:

                    st.link_button(
                        "👁 Abrir",
                        row["PDF"]
                    )

    except Exception as e:

        st.error(e)

# =========================
# INVENTARIOS
# =========================

if modulo == "📦 Inventarios":

    st.markdown("---")

    c1, c2 = st.columns(2)

    with c1:

        entidad_inv = st.selectbox(
            "📍 Entidad inventario",
            entidades
        )

    with c2:

        clues_inv = st.selectbox(
            "🏥 CLUES inventario",
            sorted(
                base_operativa[
                    base_operativa["ENTIDAD"]
                    == entidad_inv
                ]["CLUES"]
                .dropna()
                .astype(str)
                .unique()
            )
        )

    resultado_inv = base_operativa[
        base_operativa["CLUES"]
        .astype(str)
        == str(clues_inv)
    ]

    if resultado_inv.empty:

        almacen_inv = "SIN ALMACÉN"

    else:

        almacen_inv = (
            resultado_inv[
                "ALMACÉN"
            ]
            .astype(str)
            .iloc[0]
        )

    st.info(
        f"🏬 ALMACÉN: {almacen_inv}"
    )
    try:

        historial_inv = (
            inventarios_base.copy()
        )

    except:

        historial_inv = pd.DataFrame()

    inventarios_clues = historial_inv[
        historial_inv["CLUES"]
        .astype(str)
        == str(clues_inv)
    ]
    st.markdown(
        "## 📦 Evidencias inventario"
    )

    # =========================
    # INVENTARIO FÍSICO
    # =========================

    evidencia_existente = (
        inventarios_clues[
            inventarios_clues["Tipo"]
            .astype(str)
            == "Evidencia física"
        ]
    )

    if not evidencia_existente.empty:

        actual_evidencia = (
            evidencia_existente.iloc[-1]
        )

        st.success(
            "✅ Evidencia física cargada"
        )

        st.info(
            f"📷 "
            f"{actual_evidencia['Archivo']}"
        )

        st.link_button(
            "👁 Abrir actual",
            actual_evidencia["Link"],
            key="abrir_evidencia"
        )

    else:

        st.warning(
            "❌ Sin evidencia física"
        )

    evidencia_fisica = None

    if not evidencia_existente.empty:

        reemplazar_evidencia = st.checkbox(
            "♻ Sustituir evidencia física"
        )

        if reemplazar_evidencia:

            inventario_fisico = True

            evidencia_fisica = st.file_uploader(
                "📷 Nueva evidencia física",
                type=["jpg", "jpeg", "png", "pdf"],
                key="nueva_evidencia"
            )

        else:

            inventario_fisico = False

    else:

        inventario_fisico = st.checkbox(
            "☑ Inventario físico recibido"
        )

        if inventario_fisico:

            evidencia_fisica = st.file_uploader(
                "📷 Evidencia física",
                type=["jpg", "jpeg", "png", "pdf"],
                key="evidencia_fisica"
            )

    # =========================
    # ACTA INICIO
    # =========================

    acta_inicio_existente = (
        inventarios_clues[
            inventarios_clues["Tipo"]
            .astype(str)
            == "Acta de inicio"
        ]
    )

    if not acta_inicio_existente.empty:

        actual_inicio = (
            acta_inicio_existente.iloc[-1]
        )

        st.caption(
            "✅ Acta de inicio cargada"
        )

        st.info(
            f"📄 "
            f"{actual_inicio['Archivo']}"
        )

        st.link_button(
            "👁 Abrir actual",
            actual_inicio["Link"],
            key="abrir_inicio"
        )

    else:

        st.caption(
            "❌ Sin acta de inicio"
        )

    archivo_acta_inicio = None

    if not acta_inicio_existente.empty:

        reemplazar_inicio = st.checkbox(
            "♻ Sustituir acta inicio"
        )

        if reemplazar_inicio:

            archivo_acta_inicio = st.file_uploader(
                "📄 Nuevo PDF acta inicio",
                type=["pdf"],
                key="nuevo_inicio"
            )

    else:

        acta_inicio = st.checkbox(
            "☑ Acta de inicio"
        )

        if acta_inicio:

            archivo_acta_inicio = st.file_uploader(
                "📄 PDF acta inicio",
                type=["pdf"],
                key="acta_inicio"
            )
    # =========================
    # ACTA CONCLUSIÓN
    # =========================

    acta_conclusion_existente = (
        inventarios_clues[
            inventarios_clues["Tipo"]
            .astype(str)
            == "Acta de conclusión"
        ]
    )

    if not acta_conclusion_existente.empty:

        actual_conclusion = (
            acta_conclusion_existente.iloc[-1]
        )

        st.caption(
            "✅ Acta de conclusión cargada"
        )

        st.info(
            f"📄 "
            f"{actual_conclusion['Archivo']}"
        )

        st.link_button(
            "👁 Abrir actual",
            actual_conclusion["Link"],
            key="abrir_conclusion"
        )

    else:

        st.caption(
            "❌ Sin acta de conclusión"
        )
    archivo_acta_conclusion = None

    if not acta_conclusion_existente.empty:

        reemplazar_conclusion = st.checkbox(
            "♻ Sustituir acta conclusión"
        )

        if reemplazar_conclusion:

            archivo_acta_conclusion = st.file_uploader(
                "📄 Nuevo PDF acta conclusión",
                type=["pdf"],
                key="nuevo_conclusion"
            )

    else:

        acta_conclusion = st.checkbox(
            "☑ Acta de conclusión"
        )

        if acta_conclusion:

            archivo_acta_conclusion = st.file_uploader(
                "📄 PDF acta conclusión",
                type=["pdf"],
                key="acta_conclusion"
            )

    # =========================
    # INVENTARIO PDF
    # =========================

    inventario_pdf_existente = (
        inventarios_clues[
            inventarios_clues["Tipo"]
            .astype(str)
            == "Inventario PDF"
        ]
    )

    if not inventario_pdf_existente.empty:

        actual_pdf = (
            inventario_pdf_existente.iloc[-1]
        )

        st.caption(
            "✅ Inventario PDF cargado"
        )

        st.info(
            f"📄 "
            f"{actual_pdf['Archivo']}"
        )

        st.link_button(
            "👁 Abrir actual",
            actual_pdf["Link"],
            key="abrir_pdf"
        )

    else:

        st.caption(
            "❌ Sin inventario PDF"
        )

    archivo_inventario_pdf = None

    if not inventario_pdf_existente.empty:

        reemplazar_pdf = st.checkbox(
            "♻ Sustituir inventario PDF"
        )

        if reemplazar_pdf:

            archivo_inventario_pdf = st.file_uploader(
                "📄 Nuevo inventario PDF",
                type=["pdf"],
                key="nuevo_pdf"
            )

    else:

        inventario_pdf = st.checkbox(
            "☑ Inventario PDF"
        )

        if inventario_pdf:

            archivo_inventario_pdf = st.file_uploader(
                "📄 Archivo inventario PDF",
                type=["pdf"],
                key="inventario_pdf"
            )

    # =========================
    # INVENTARIO EXCEL
    # =========================

    inventario_excel_existente = (
        inventarios_clues[
            inventarios_clues["Tipo"]
            .astype(str)
            == "Inventario Excel"
        ]
    )

    if not inventario_excel_existente.empty:

        actual_excel = (
            inventario_excel_existente.iloc[-1]
        )

        st.caption(
            "✅ Inventario Excel cargado"
        )

        st.info(
            f"📊 "
            f"{actual_excel['Archivo']}"
        )

        st.link_button(
            "👁 Abrir actual",
            actual_excel["Link"],
            key="abrir_excel"
        )

    else:

        st.caption(
            "❌ Sin inventario Excel"
        )

    archivo_inventario_excel = None

    if not inventario_excel_existente.empty:

        reemplazar_excel = st.checkbox(
            "♻ Sustituir inventario Excel"
        )

        if reemplazar_excel:

            archivo_inventario_excel = st.file_uploader(
                "📊 Nuevo inventario Excel",
                type=["xlsx", "xls"],
                key="nuevo_excel"
            )

    else:

        inventario_excel = st.checkbox(
            "☑ Inventario Excel"
        )

        if inventario_excel:

            archivo_inventario_excel = st.file_uploader(
                "📊 Archivo inventario Excel",
                type=["xlsx", "xls"],
                key="inventario_excel"
            )

    if st.button("📤 Guardar inventario"):
        archivos_subidos = []

        if evidencia_fisica:

            archivos_subidos.append(
                (
                    "Evidencia física",
                    evidencia_fisica
                )
            )

        if archivo_acta_inicio:

            archivos_subidos.append(
                (
                    "Acta de inicio",
                    archivo_acta_inicio
                )
            )

        if archivo_acta_conclusion:

            archivos_subidos.append(
                (
                    "Acta de conclusión",
                    archivo_acta_conclusion
                )
            )

        if archivo_inventario_pdf:

            archivos_subidos.append(
                (
                    "Inventario PDF",
                    archivo_inventario_pdf
                )
            )

        if archivo_inventario_excel:

            archivos_subidos.append(
                (
                    "Inventario Excel",
                    archivo_inventario_excel
                )
            )

        if not archivos_subidos:

            st.warning(
                "⚠ Debes subir al menos un archivo"
            )

        else:

            folder_id = (
                obtener_carpeta_inventarios(
                    entidad_inv,
                    clues_inv
                )
            )

            for tipo_archivo, archivo_actual in archivos_subidos:
                historial_inv = (
                    inventarios_base.copy()
                )

                existente = historial_inv[
                    (
                        historial_inv["CLUES"]
                        .astype(str)
                        == str(clues_inv)
                    )
                    &
                    (
                        historial_inv["Tipo"]
                        .astype(str)
                        == str(tipo_archivo)
                    )
                ]
                if not existente.empty:

                    archivo_existente = (
                        existente.iloc[-1]
                    )

                    try:

                        drive_service.files().update(
                            fileId=archivo_existente["File ID"],
                            body={
                                "trashed": True
                            },
                            supportsAllDrives=True
                        ).execute()

                    except Exception as e:

                        st.error(e)

                        continue
                with tempfile.NamedTemporaryFile(
                    delete=False
                ) as temp_file:

                    temp_file.write(
                        archivo_actual.getbuffer()
                    )

                    temp_path = temp_file.name

                nombre_drive = (
                    f"{tipo_archivo}_{clues_inv}_"
                    f"{archivo_actual.name}"
                )

                file_metadata = {
                    "name": nombre_drive,
                    "parents": [folder_id]
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

                os.remove(temp_path)

                guardar_inventario_sheets(
                    fecha=pd.Timestamp.now(),
                    entidad=entidad_inv,
                    clues=clues_inv,
                    tipo=tipo_archivo,
                    archivo=archivo_actual.name,
                    inventario_fisico=(
                        "SI"
                        if inventario_fisico
                        else "NO"
                    ),
                    link=uploaded_file[
                        "webViewLink"
                    ],
                    file_id=uploaded_file["id"]
                )

            st.cache_data.clear()

            st.toast(
                "✅ Inventarios guardados",
                icon="✅"
            )

            st.rerun()
    st.markdown("---")

    st.markdown(
        "## 📦 Historial inventarios"
    )

    try:

        historial_inv = inventarios_base.copy()

        historial_tabla = historial_inv[
            [
                "Fecha",
                "Entidad",
                "CLUES",
                "Tipo",
                "Archivo",
                "Inventario físico"
            ]
        ].copy()

        st.dataframe(

            historial_tabla,

            use_container_width=True,

            hide_index=True,

            height=500
        )

    except Exception as e:

        st.error(e)