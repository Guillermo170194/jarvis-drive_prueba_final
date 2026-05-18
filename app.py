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

df_incorrectos = base_operativa[
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

resumen_entidad = []

for entidad_nombre in sorted(
    base_operativa["ENTIDAD"]
    .dropna()
    .astype(str)
    .unique()
):

    historial_entidad = historial_docs[
        historial_docs["Entidad"]
        .astype(str)
        == entidad_nombre
    ]

    entrega = historial_entidad[
        historial_entidad["Tipo"]
        == "Entrega"
    ].shape[0]

    entrega_uas = historial_entidad[
        historial_entidad["Tipo"]
        == "Entrega UAS Y OIC"
    ].shape[0]

    correccion = historial_entidad[
        historial_entidad["Tipo"]
        == "Corrección"
    ].shape[0]

    reiterativo_1 = historial_entidad[
        historial_entidad["Tipo"]
        == "Primer reiterativo"
    ].shape[0]

    reiterativo_2 = historial_entidad[
        historial_entidad["Tipo"]
        == "Segundo reiterativo"
    ].shape[0]

    reiterativo_3 = historial_entidad[
        historial_entidad["Tipo"]
        == "Tercer reiterativo"
    ].shape[0]
    prorroga = historial_entidad[
        historial_entidad["Tipo"]
        == "Prórroga"
    ].shape[0]

    correo = historial_entidad[
        historial_entidad["Tipo"]
        == "Correo"
    ].shape[0]

    resumen_entidad.append({

        "Entidad": entidad_nombre,

        "Entrega": entrega,

        "Entrega UAS Y OIC": entrega_uas,

        "Corrección": correccion,

        "1er Reiterativo": reiterativo_1,

        "2do Reiterativo": reiterativo_2,

        "3er Reiterativo": reiterativo_3,

        "Prórroga": prorroga,

        "Correo": correo

    })

df_resumen_entidad = pd.DataFrame(
    resumen_entidad
)

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
    st.markdown("---")

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

    st.markdown(
        "## 📊 Resumen nacional por entidad"
    )

    st.dataframe(
        df_resumen_entidad,
        use_container_width=True,
        hide_index=True
    )
    st.markdown("---")

    c1, c2, c3 = st.columns(3)

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
            "## ❌ Incorrectos"
        )

        st.dataframe(
            df_incorrectos[
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

    with c3:

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