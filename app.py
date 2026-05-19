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
# DESCARGAR EXCEL
# =========================

@st.cache_data(ttl=30)
def descargar_base_operativa():

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

    while done is False:

        status, done = downloader.next_chunk()

    archivo.seek(0)

    return pd.read_excel(
        archivo,
        sheet_name="HISTORIAL_DOCUMENTAL"
    )


@st.cache_data(ttl=30)
def descargar_inventarios():

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

    while done is False:

        status, done = downloader.next_chunk()

    archivo.seek(0)

    return pd.read_excel(
        archivo,
        sheet_name="INVENTARIOS"
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
        range="HISTORIAL_DOCUMENTAL!A:H"
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
# KPIs NACIONALES
# =========================

inventarios_base = descargar_inventarios()

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

historial_docs = descargar_historial()

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

inventarios_base = descargar_inventarios()

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

    fila = {
        "Entidad": entidad,
        "CLUES": clues,
        "ALMACÉN": almacen,
        "INVENTARIO FÍSICO": (
            "SI"
            if clues in clues_con_inventario
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
    almacen_clues = (
        base_operativa[
            base_operativa["CLUES"]
            .astype(str)
            == str(clues)
        ]["ALMACÉN"]
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
                            range="HISTORIAL_DOCUMENTAL!A:H"
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

# =========================
# SUPERVISIÓN
# =========================

if modulo == "🕵 Supervisión":

    st.markdown("---")

    st.markdown(
        "# 🕵 Cédula de supervisión"
    )

    # =========================
    # DATOS GENERALES
    # =========================

    s1, s2 = st.columns(2)

    with s1:

        entidad_sup = st.selectbox(
            "📍 Entidad",
            entidades,
            key="sup_entidad"
        )

    with s2:

        clues_sup = st.selectbox(
            "🏥 CLUES",
            sorted(
                base_operativa[
                    base_operativa["ENTIDAD"]
                    == entidad_sup
                ]["CLUES"]
                .dropna()
                .astype(str)
                .unique()
            ),
            key="sup_clues"
        )

    almacen_sup = (
        base_operativa[
            base_operativa["CLUES"]
            .astype(str)
            == str(clues_sup)
        ]["ALMACÉN"]
        .astype(str)
        .iloc[0]
    )

    fecha_supervision = st.date_input(
        "📅 Fecha supervisión"
    )

    st.info(
        f"🏬 ALMACÉN: {almacen_sup}"
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
        "# 📋 Diagnóstico general"
    )

    conceptos_generales = [
        "ANEXO 1.- LISTADO GRAL",
        "ANEXO 2.- IMSS B",
        "ANEXO 3.- PROPIOS",
        "Acta de conclusión",
        "ANEXO 5.- ACTA DE INICIO"
    ]

    for concepto in conceptos_generales:

        st.markdown(f"### {concepto}")

        c1, c2, c3, c4 = st.columns(4)

        with c1:

            contiene = st.selectbox(
                "Contiene",
                ["SI", "NO"],
                key=f"{concepto}_contiene"
            )

        with c2:

            piezas = st.number_input(
                "Piezas",
                min_value=0.0,
                step=1.0,
                key=f"{concepto}_piezas"
            )

        with c3:

            monto = st.number_input(
                "Monto",
                min_value=0.0,
                step=1.0,
                key=f"{concepto}_monto"
            )

        with c4:

            firmado = st.selectbox(
                "Firmado",
                ["SI", "NO"],
                key=f"{concepto}_firmado"
            )

        observaciones = st.text_area(
            "Observaciones",
            key=f"{concepto}_obs"
        )

        st.markdown("---")

    # =========================
    # DIFERENCIAS
    # =========================

    st.markdown(
        "# ⚠ Diferencias"
    )

    conceptos_diferencias = [
        "ANEXO 4.- REPORTE DE DIF",
        "ANEXO 6.- REPORTE LENTO Y NULO",
        "ANEXO 7.- REPORTE CADUCOS",
        "ANEXO 8.- REP DE PROX A CADUCAR",
        "EXCEL",
        "PDF",
        "FISICO"
    ]

    for concepto in conceptos_diferencias:

        st.markdown(f"### {concepto}")

        d1, d2, d3 = st.columns(3)

        with d1:

            existe = st.selectbox(
                "Existe",
                ["SI", "NO"],
                key=f"{concepto}_existe"
            )

        with d2:

            dif_mas = st.number_input(
                "Dif más",
                step=1.0,
                key=f"{concepto}_mas"
            )

        with d3:

            dif_menos = st.number_input(
                "Dif menos",
                step=1.0,
                key=f"{concepto}_menos"
            )

        observaciones = st.text_area(
            "Observaciones",
            key=f"{concepto}_obs_2"
        )

        st.markdown("---")

    # =========================
    # FIRMAS
    # =========================

    st.markdown(
        "# ✍ Firmas"
    )

    f1, f2 = st.columns(2)

    with f1:

        firma_verificador = st.file_uploader(
            "Firma verificador",
            type=["png", "jpg", "jpeg"],
            key="firma_verificador"
        )

    with f2:

        firma_almacen = st.file_uploader(
            "Firma almacén",
            type=["png", "jpg", "jpeg"],
            key="firma_almacen"
        )

    st.markdown("---")

    if st.button("📤 Generar cédula"):

        st.success(
            "✅ Estructura supervisión lista"
        )

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

    almacen_inv = (
        base_operativa[
            base_operativa["CLUES"]
            .astype(str)
            == str(clues_inv)
        ]["ALMACÉN"]
        .astype(str)
        .iloc[0]
    )

    st.info(
        f"🏬 ALMACÉN: {almacen_inv}"
    )
    try:

        historial_inv = (
            descargar_inventarios()
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
                    descargar_inventarios()
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

            st.success(
                "✅ Inventarios guardados"
            )
    st.markdown("---")

    st.markdown(
        "## 📦 Historial inventarios"
    )

    try:

        historial_inv = descargar_inventarios()

        for i, row in historial_inv.iterrows():

            c1, c2, c3, c4, c5 = st.columns(
                [2, 2, 3, 1, 1]
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
                    row["Link"],
                    key=f"abrir_inv_{i}"
                )

            with c5:

                st.write(
                    row[
                        "Inventario físico"
                    ]
                )

    except Exception as e:

        st.error(e)