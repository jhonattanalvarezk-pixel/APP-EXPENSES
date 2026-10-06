from datetime import datetime
import json
import os
import re
import tempfile
import time
from google import genai
from google.genai import types
import openpyxl
import pandas as pd
import json
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st
googleapiclient.discovery

# Define los scopes de Google Drive al inicio del archivo o aquí mismo
SCOPES = ["https://www.googleapis.com/auth/drive"]


def get_drive_service():
  service_account_info = {
      "type": st.secrets["type"],
      "project_id": st.secrets["project_id"],
      "private_key_id": st.secrets["private_key_id"],
      "private_key": st.secrets["private_key"].replace("\\n", "\n"),
      "client_email": st.secrets["client_email"],
      "client_id": st.secrets["client_id"],
      "auth_uri": st.secrets["auth_uri"],
      "token_uri": st.secrets["token_uri"],
      "auth_provider_x509_cert_url": st.secrets["auth_provider_x509_cert_url"],
      "client_x509_cert_url": st.secrets["client_x509_cert_url"],
      "universe_domain": st.secrets["universe_domain"],
  }

  creds = ServiceAccountCredentials.from_json_keyfile_dict(
      service_account_info, SCOPES
  )
  service = build("drive", "v3", credentials=creds)
  return service
def list_files(service):
        results = service.files().list(pageSize=10, fields="files(id, name)", q="'1dmpCWssJGY295gx-h5V90xmbHLRdrUNE' in parents").execute()
        files = results.get("files", [])
        if files:
            st.write("Archivos:")
        for file in files:
            st.write(f"{file['name']} ({file['id']})")

# Configuración de la página
st.set_page_config(
    page_title="Gestión de Facturas y Gastos Pro", page_icon="🧾", layout="wide"
)

# Estilos CSS personalizados para interfaz elegante
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0F172A;
        color: #F8FAFC;
    }
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
        border-right: 1px solid #334155;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        color: #38BDF8 !important;
    }
    div[data-testid="metric-container"] {
        background-color: #1E293B;
        border: 1px solid #334155;
        padding: 15px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white;
        border-radius: 8px;
        font-weight: 600;
        border: none;
        padding: 0.6rem 1.2rem;
    }
    .stForm {
        background-color: #1E293B;
        border: 1px solid #334155 !important;
        padding: 20px;
        border-radius: 12px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- MOSTRAR LOGO EN LA BARRA LATERAL ---
LOGO_PATH = "logo.png"  # Cambia esto por la ruta o nombre de tu archivo de logo

if os.path.exists(LOGO_PATH):
  st.sidebar.image(LOGO_PATH, use_container_width=True)
else:
  st.sidebar.info(f"💡 Coloca tu archivo '{LOGO_PATH}' en la carpeta del app.")

# Configuración de API Key
st.sidebar.title("Configuracion")
api_key = st.sidebar.text_input(
    "Gemini API Key", type="password", value=os.environ.get("GEMINI_API_KEY", "")
)

EXCEL_FILE = "registro_facturas.xlsx"
PDF_DIR = "archivos_facturas"

os.makedirs(PDF_DIR, exist_ok=True)


def inicializar_excel():
  if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"
    headers = [
        "ID",
        "Fecha",
        "Proveedor",
        "Origen",
        "Categoria",
        "Base Imponible",
        "IVA (%)",
        "Cuota IVA",
        "Total",
        "Archivo Original",
    ]
    ws.append(headers)
    wb.save(EXCEL_FILE)


inicializar_excel()


def limpiar_json_string(texto: str) -> str:
  texto = re.sub(r"^```json\s*", "", texto, flags=re.MULTILINE)
  texto = re.sub(r"^```\s*", "", texto, flags=re.MULTILINE)
  texto = re.sub(r"```$", "", texto, flags=re.MULTILINE)
  return texto.strip()


def extraer_datos_factura(file_bytes, mime_type, api_key):
  client = genai.Client(api_key=api_key)

  prompt = """
    Analiza este documento de factura o ticket y extrae los siguientes datos en formato JSON estricto:
    - fecha: Fecha de la factura en formato AAAA-MM-DD (si no se encuentra, pon la fecha actual).
    - proveedor: Nombre de la tienda, empresa o emisor de la factura.
    - categoria: Clasifica el gasto en una de estas opciones: Materiales, Transporte, Herramientas, Varios.
    - base_imponible: Numero decimal con la base imponible total.
    - iva_porcentaje: Numero entero o decimal con el tipo de IVA principal (ej: 21, 10, 4).
    - cuota_iva: Numero decimal con el importe total del IVA.
    - total: Numero decimal con el importe total de la factura.
    
    Devuelve UNICAMENTE un JSON valido con estas claves: fecha, proveedor, categoria, base_imponible, iva_porcentaje, cuota_iva, total.
    """

  ext = ".pdf" if "pdf" in mime_type else ".jpg"

  with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_file:
    tmp_file.write(file_bytes)
    tmp_path = tmp_file.name

  uploaded_file = None
  try:
    uploaded_file = client.files.upload(file=tmp_path)

    if "pdf" in mime_type:
      while uploaded_file.state.name == "PROCESSING":
        time.sleep(1)
        uploaded_file = client.files.get(name=uploaded_file.name)

      if uploaded_file.state.name == "FAILED":
        raise Exception("Google Gemini no pudo procesar este archivo PDF.")

    modelos = ["gemini-1.5-flash",]
    ultimo_error = None

    for modelo in modelos:
      for intento in range(4):
        try:
          response = client.models.generate_content(
              model=modelo,
              contents=[uploaded_file, prompt],
              config=types.GenerateContentConfig(
                  response_mime_type="application/json"
              ),
          )
          return response.text
        except Exception as e:
          ultimo_error = e
          msg_error = str(e)

          if "404" in msg_error or "NOT_FOUND" in msg_error:
            break

          if (
              "503" in msg_error
              or "UNAVAILABLE" in msg_error
              or "RESOURCE_EXHAUSTED" in msg_error
              or "429" in msg_error
          ):
            time.sleep(2**intento)
            continue
          else:
            break

    if ultimo_error:
      raise ultimo_error

  finally:
    if os.path.exists(tmp_path):
      os.remove(tmp_path)
    if uploaded_file:
      try:
        client.files.delete(name=uploaded_file.name)
      except Exception:
        pass

  raise Exception(
      "Los servidores de Gemini estan con alta demanda. Reintentalo en unos"
      " instantes."
  )


# --- CABECERA PRINCIPAL ---
col_head1, col_head2 = st.columns([0.8, 0.2])
with col_head1:
  st.title("Gestor de Facturas y Gastos Pro")
with col_head2:
  if os.path.exists(LOGO_PATH):
    st.image(LOGO_PATH, width=120)

menu = st.sidebar.selectbox(
    "Menu",
    [
        "Nueva Factura (Individual/Lote)",
        "Listado y Edicion Interactiva",
        "Resumen Trimestral y Exportacion",
    ],
)

# 1. OPCION: NUEVA FACTURA
if menu == "Nueva Factura (Individual/Lote)":
  st.header("Carga de Facturas y Tickets")

  origen_seleccionado = st.radio(
      "De donde vienen las facturas?",
      ["Manual / Camara", "WhatsApp / PDF Masivo"],
      horizontal=True,
  )

  uploaded_files = st.file_uploader(
      "Sube una o varias imagenes o PDFs",
      type=["png", "jpg", "jpeg", "pdf"],
      accept_multiple_files=True,
  )

  if uploaded_files and not api_key:
    st.warning("Por favor, introduce tu Gemini API Key en la barra lateral.")

  elif uploaded_files and api_key:
    if st.button("Procesar Facturas con IA", type="primary"):
      for uploaded_file in uploaded_files:
        with st.spinner(f"Procesando {uploaded_file.name}..."):
          file_bytes = uploaded_file.getvalue()
          mime_type = uploaded_file.type

          timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
          safe_filename = f"{timestamp}_{uploaded_file.name}"
          file_path = os.path.join(PDF_DIR, safe_filename)
          with open(file_path, "wb") as f:
            f.write(file_bytes)

          try:
            json_str = extraer_datos_factura(file_bytes, mime_type, api_key)
            json_str_limpio = limpiar_json_string(json_str)
            datos = json.loads(json_str_limpio)

            st.session_state[f"datos_{safe_filename}"] = {
                "datos": datos,
                "safe_filename": safe_filename,
                "file_path": file_path,
                "mime_type": mime_type,
            }
            st.success(f"Procesado con exito: {uploaded_file.name}!")

          except Exception as e:
            st.error(f"Error procesando {uploaded_file.name}: {e}")

  # Renderizado de pendientes
  keys_to_show = [
      k for k in st.session_state.keys() if k.startswith("datos_")
  ]

  if keys_to_show:
    st.subheader("Facturas pendientes de guardar en Excel")

    if len(keys_to_show) > 1:
      if st.button(
          f"Guardar TODAS ({len(keys_to_show)}) en Excel de golpe",
          type="secondary",
      ):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active

        for key in list(keys_to_show):
          item = st.session_state[key]
          datos = item["datos"]
          safe_filename = item["safe_filename"]

          nuevo_id = ws.max_row
          origen_etiqueta = (
              "WhatsApp"
              if origen_seleccionado == "WhatsApp / PDF Masivo"
              else "Manual/Camara"
          )

          ws.append([
              nuevo_id,
              str(datos.get("fecha", datetime.now().strftime("%Y-%m-%d"))),
              str(datos.get("proveedor", "Desconocido")),
              origen_etiqueta,
              str(datos.get("categoria", "Materiales")),
              float(datos.get("base_imponible") or 0.00),
              float(datos.get("iva_porcentaje") or 21.0),
              float(datos.get("cuota_iva") or 0.00),
              float(datos.get("total") or 0.00),
              safe_filename,
          ])
          del st.session_state[key]

        wb.save(EXCEL_FILE)
        st.success("Todas las facturas del lote han sido guardadas en Excel!")
        time.sleep(1)
        st.rerun()

    for key in keys_to_show:
      item = st.session_state[key]
      datos = item["datos"]
      safe_filename = item["safe_filename"]
      file_path = item["file_path"]
      mime_type = item["mime_type"]

      fecha_str = datos.get("fecha", "")
      try:
        fecha_valida = datetime.strptime(fecha_str, "%Y-%m-%d").date()
      except (ValueError, TypeError):
        fecha_valida = datetime.now().date()

      with st.expander(
          f"Revisar: {safe_filename.split('_', 2)[-1]}", expanded=True
      ):
        col_form, col_visor = st.columns([1.2, 0.8])

        with col_visor:
          st.markdown("**Visor del documento**")
          if "pdf" in mime_type:
            st.info(f"PDF adjunto guardado en: {safe_filename}")
          else:
            st.image(file_path, use_container_width=True)

        with col_form:
          with st.form(key=f"form_{safe_filename}"):
            f_fecha = st.date_input("Fecha", value=fecha_valida)
            f_proveedor = st.text_input(
                "Proveedor", value=str(datos.get("proveedor", ""))
            )
            cat_detectada = str(datos.get("categoria", "Materiales"))
            categorias_validas = [
                "Materiales",
                "Transporte",
                "Herramientas",
                "Varios",
            ]
            index_cat = (
                categorias_validas.index(cat_detectada)
                if cat_detectada in categorias_validas
                else 0
            )
            f_categoria = st.selectbox(
                "Categoria", categorias_validas, index=index_cat
            )

            f_base = st.number_input(
                "Base Imponible (EUR)",
                value=float(datos.get("base_imponible") or 0.00),
                format="%.2f",
            )
            f_iva_p = st.number_input(
                "IVA (%)",
                value=float(datos.get("iva_porcentaje") or 21.0),
                format="%.1f",
            )
            f_cuota = st.number_input(
                "Cuota IVA (EUR)",
                value=float(datos.get("cuota_iva") or 0.00),
                format="%.2f",
            )
            f_total = st.number_input(
                "Total (EUR)",
                value=float(datos.get("total") or 0.00),
                format="%.2f",
            )

            submitted = st.form_submit_button("Guardar Factura en Excel")
            if submitted:
              wb = openpyxl.load_workbook(EXCEL_FILE)
              ws = wb.active
              nuevo_id = ws.max_row
              origen_etiqueta = (
                  "WhatsApp"
                  if origen_seleccionado == "WhatsApp / PDF Masivo"
                  else "Manual/Camara"
              )

              ws.append([
                  nuevo_id,
                  str(f_fecha),
                  f_proveedor,
                  origen_etiqueta,
                  f_categoria,
                  f_base,
                  f_iva_p,
                  f_cuota,
                  f_total,
                  safe_filename,
              ])
              wb.save(EXCEL_FILE)

              del st.session_state[key]
              st.success(
                  f"Factura de {f_proveedor} guardada correctamente en Excel."
              )
              time.sleep(1)
              st.rerun()

# 2. OPCION: LISTADO Y EDICION INTERACTIVA
elif menu == "Listado y Edicion Interactiva":
  st.header("Listado, Edicion y Registro de Facturas")
  if os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)
    if not df.empty:
      st.markdown(
          "Puedes **editar valores directamente en la tabla** o seleccionar las"
          " casillas para guardar o eliminar filas."
      )

      edited_df = st.data_editor(
          df,
          num_rows="dynamic",
          use_container_width=True,
          key="tabla_interactiva",
      )

      if st.button("Guardar Cambios Realizados en la Tabla"):
        edited_df.to_excel(EXCEL_FILE, index=False)
        st.success("El archivo Excel ha sido actualizado con los cambios!")
        time.sleep(1)
        st.rerun()

      st.divider()
      st.subheader("Visor de Archivo Guardado")
      archivos_disponibles = [
          f for f in df["Archivo Original"].dropna().unique()
      ]
      if archivos_disponibles:
        archivo_sel = st.selectbox(
            "Selecciona una factura para consultar el documento:",
            archivos_disponibles,
        )
        ruta_archivo = os.path.join(PDF_DIR, archivo_sel)
        if os.path.exists(ruta_archivo):
          if archivo_sel.lower().endswith(".pdf"):
            st.info(f"El archivo es un PDF guardado en: {ruta_archivo}")
          else:
            st.image(ruta_archivo, width=500)
        else:
          st.warning("El archivo original no se encuentra en el servidor.")
    else:
      st.info("No hay facturas registradas todavia.")
  else:
    st.info("El archivo Excel aun no ha sido creado.")

# 3. OPCION: RESUMEN TRIMESTRAL Y EXPORTACION COMPLETA
elif menu == "Resumen Trimestral y Exportacion":
  st.header("Resumen Trimestral, IVA Soportado y Exportacion")
  if os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)
    if not df.empty:
      total_base = df["Base Imponible"].sum()
      total_iva = df["Cuota IVA"].sum()
      total_general = df["Total"].sum()

      col1, col2, col3 = st.columns(3)
      col1.metric("Base Imponible Total", f"{total_base:.2f} EUR")
      col2.metric("IVA Soportado Total", f"{total_iva:.2f} EUR")
      col3.metric("Gasto Total", f"{total_general:.2f} EUR")

      st.subheader("Desglose por Categoria")
      col_cat = "Categoria" if "Categoria" in df.columns else "Categoría"
      resumen_cat = (
          df.groupby(col_cat)[["Base Imponible", "Cuota IVA", "Total"]]
          .sum()
          .reset_index()
      )
      st.dataframe(resumen_cat, use_container_width=True)

      st.divider()
      st.subheader("Exportacion de Informes")

      with open(EXCEL_FILE, "rb") as f:
        excel_bytes = f.read()

      st.download_button(
          label="Descargar Excel de Registro de Facturas (.xlsx)",
          data=excel_bytes,
          file_name=f"registro_facturas_{datetime.now().strftime('%Y%m%d')}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )
    else:
      st.info("Aun no hay datos para resumir.")
  else:
    st.info("No hay registros disponibles.")
service=get_drive_service()
list_files(service)

