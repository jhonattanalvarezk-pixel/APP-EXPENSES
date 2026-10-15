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
import io
import streamlit as st
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from PIL import Image, ImageEnhance
import streamlit.components.v1 as components

def get_drive_service():
    creds = Credentials(
        None,
        refresh_token=st.secrets["OAUTH_REFRESH_TOKEN"],
        client_id=st.secrets["OAUTH_CLIENT_ID"],
        client_secret=st.secrets["OAUTH_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
    )
    if not creds.valid:
        creds.refresh(Request())
    return build("drive", "v3", credentials=creds)

def subir_archivo_a_drive(file_path, file_name):
    try:
        service = get_drive_service()
        folder_id = st.secrets.get('DRIVE_FOLDER_ID')
        query = f"name = '{file_name}' and '{folder_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id)").execute()
        files = results.get("files", [])
        media = MediaFileUpload(file_path, resumable=True)
        if files:
            file_id = files[0]['id']
            file = service.files().update(fileId=file_id, media_body=media).execute()
        else:
            file_metadata = {'name': file_name, 'parents': [folder_id]}
            file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return file.get('id')
    except Exception as e:
        print(f"Error Drive: {e}")
        return None

def descargar_archivo_de_drive(nombre_archivo, destino_path):
    try:
        service = get_drive_service()
        folder_id = st.secrets.get('DRIVE_FOLDER_ID')
        query = f"name = '{nombre_archivo}' and '{folder_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get("files", [])
        if not files:
            return False
        file_id = files[0]['id']
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        with open(destino_path, "wb") as f:
            f.write(fh.read())
        return True
    except Exception as e:
        print(f"Error descargando: {e}")
        return False

def mostrar_archivo_inteligente(file_path, archivo_nombre):
    if not os.path.exists(file_path):
        with st.spinner("Recuperando documento..."):
            encontrado = descargar_archivo_de_drive(archivo_nombre, file_path)
            if not encontrado:
                st.error("El archivo no se encuentra disponible localmente.")
                return
    if archivo_nombre.lower().endswith(".pdf"):
        try:
            with open(file_path, "rb") as f:
                pdf_bytes = f.read()
            st.download_button("📥 Descargar PDF original", data=pdf_bytes, file_name=archivo_nombre, mime="application/pdf", type="primary")
            import fitz
            doc = fitz.open(file_path)
            page = doc[0]
            pix = page.get_pixmap(dpi=150)
            img_path = file_path + "_vis.png"
            pix.save(img_path)
            st.image(img_path, use_container_width=True)
        except Exception as e:
            st.error(f"Error en PDF: {e}")
    else:
        st.image(file_path, use_container_width=True)

def aclarar_y_optimizar_imagen(image_bytes):
    try:
        img = Image.open(io.BytesIO(image_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        max_width = 1200
        if img.width > max_width:
            w_percent = (max_width / float(img.width))
            h_size = int(float(img.height) * float(w_percent))
            img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
        img = ImageEnhance.Contrast(img).enhance(1.5)
        img = ImageEnhance.Sharpness(img).enhance(1.6)
        output_io = io.BytesIO()
        img.save(output_io, format="JPEG", quality=85, optimize=True)
        return output_io.getvalue()
    except Exception:
        return image_bytes

st.set_page_config(page_title="Gestión de Facturas Pro", page_icon="🧾", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #0F172A; color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #1E293B !important; }
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white; border-radius: 8px; font-weight: 600; border: none; width: 100%;
    }
    .stForm { background-color: #1E293B; border: 1px solid #334155 !important; border-radius: 12px; }
    [data-testid="stImage"] img { border-radius: 10px; max-height: 380px; object-fit: contain; background-color: #0B0F19; }
    </style>
""", unsafe_allow_html=True)

LOGO_PATH = "logo.png"
if os.path.exists(LOGO_PATH):
  st.sidebar.image(LOGO_PATH, use_container_width=True)

st.sidebar.title("Configuración")
api_key = st.sidebar.text_input("Gemini API Key", type="password", value=os.environ.get("GEMINI_API_KEY", ""))

EXCEL_FILE = "registro_facturas.xlsx"
PDF_DIR = "archivos_facturas"
os.makedirs(PDF_DIR, exist_ok=True)

def inicializar_excel():
  if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append(["ID", "Fecha", "Proveedor", "Tipo", "Origen", "Categoria", "Base Imponible", "IVA (%)", "Cuota IVA", "Total", "Archivo Original"])
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
    Analiza este documento de factura o ticket y extrae en JSON estricto:
    - fecha: AAAA-MM-DD (si no hay, fecha actual).
    - proveedor: Nombre del emisor.
    - categoria: Materiales, Transporte, Herramientas, Varios.
    - base_imponible: Número decimal.
    - iva_porcentaje: Número.
    - cuota_iva: Número decimal.
    - total: Número decimal.
    Devuelve UNICAMENTE un JSON válido con estas claves.
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

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[uploaded_file, prompt],
        config=types.GenerateContentConfig(response_mime_type="application/json"),
    )
    return response.text

  finally:
    if os.path.exists(tmp_path):
      os.remove(tmp_path)
    if uploaded_file:
      try:
        client.files.delete(name=uploaded_file.name)
      except:
        pass

col_h1, col_h2 = st.columns([0.8, 0.2])
with col_h1:
  st.title("Gestión de Facturas Pro")

menu = st.sidebar.selectbox("Menú", ["Nueva Factura (Individual/Lote)", "Listado y Edición Interactiva", "Resumen Trimestral y Exportación"])

if menu == "Nueva Factura (Individual/Lote)":
  st.header("Carga de Facturas y Tickets")

  uploaded_file = st.file_uploader("Selecciona imagen o PDF", type=["png", "jpg", "jpeg", "webp", "pdf"], key="u_unico")

  if uploaded_file:
    if not api_key:
      st.warning("⚠️ Introduce tu Gemini API Key en la barra lateral.")
    else:
      if st.button("✨ Procesar con Inteligencia Artificial", type="primary"):
        with st.spinner("Analizando documento..."):
          try:
            file_bytes = uploaded_file.getvalue()
            mime_type = uploaded_file.type

            if mime_type.startswith("image/"):
              file_bytes = aclarar_y_optimizar_imagen(file_bytes)
              mime_type = "image/jpeg"

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_filename = f"{timestamp}_{uploaded_file.name}"
            file_path = os.path.join(PDF_DIR, safe_filename)
            
            with open(file_path, "wb") as f:
              f.write(file_bytes)

            json_str = extraer_datos_factura(file_bytes, mime_type, api_key)
            datos = json.loads(limpiar_json_string(json_str))

            st.session_state["factura_actual"] = {
                "datos": datos,
                "safe_filename": safe_filename,
                "file_path": file_path,
                "file_bytes": file_bytes,
            }
            st.success("¡Procesado con éxito!")
            st.rerun()
          except Exception as e:
            st.error(f"Error al procesar: {e}")

  if "factura_actual" in st.session_state:
    item = st.session_state["factura_actual"]
    datos = item["datos"]
    safe_filename = item["safe_filename"]
    file_path = item["file_path"]
    file_bytes = item["file_bytes"]

    if not os.path.exists(file_path) and file_bytes:
      with open(file_path, "wb") as f:
        f.write(file_bytes)

    st.divider()
    st.subheader("Confirmar datos y guardar")

    col_img, col_form = st.columns([1, 1])
    with col_img:
      if os.path.exists(file_path):
        if safe_filename.lower().endswith(".pdf"):
          st.info("📄 Documento PDF")
        else:
          st.image(file_path, use_container_width=True)

    with col_form:
      with st.form("form_confirmacion"):
        tipo_comprobante = st.radio("Tipo", ["Ticket", "Factura"], horizontal=True)
        f_proveedor = st.text_input("Proveedor", value=str(datos.get("proveedor", "")))
        
        try:
          f_fecha = datetime.strptime(str(datos.get("fecha", "")), "%Y-%m-%d").date()
        except:
          f_fecha = datetime.now().date()
        f_fecha = st.date_input("Fecha", value=f_fecha)

        f_total = st.number_input("Total (€)", value=float(datos.get("total") or 0.00), format="%.2f")
        f_base = st.number_input("Base Imponible (€)", value=float(datos.get("base_imponible") or 0.00), format="%.2f")
        f_iva_p = st.number_input("IVA (%)", value=float(datos.get("iva_porcentaje") or 21.0), format="%.1f")
        f_cuota = st.number_input("Cuota IVA (€)", value=float(datos.get("cuota_iva") or 0.00), format="%.2f")

        submitted = st.form_submit_button("GUARDAR EN EXCEL")
        if submitted:
          try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active
            ws.append([
                ws.max_row, str(f_fecha), f_proveedor, tipo_comprobante, 
                "Cámara Móvil", str(datos.get("categoria", "Materiales")), 
                f_base, f_iva_p, f_cuota, f_total, safe_filename
            ])
            wb.save(EXCEL_FILE)
            subir_archivo_a_drive(file_path, safe_filename)
            subir_archivo_a_drive(EXCEL_FILE, EXCEL_FILE)

            del st.session_state["factura_actual"]
            st.success("¡Guardado con éxito!")
            time.sleep(1)
            st.rerun()
          except Exception as e:
            st.error(f"Error al guardar: {e}")

elif menu == "Listado y Edición Interactiva":
  st.header("Listado de Facturas")
  if os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)
    if not df.empty:
      edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
      if st.button("Guardar Cambios"):
        edited_df.to_excel(EXCEL_FILE, index=False)
        subir_archivo_a_drive(EXCEL_FILE, EXCEL_FILE)
        st.success("¡Actualizado!")
        st.rerun()
    else:
      st.info("No hay registros.")

elif menu == "Resumen Trimestral y Exportación":
  st.header("Resumen")
  if os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)
    if not df.empty:
      st.metric("Gasto Total", f"{df['Total'].sum():.2f} EUR")
      with open(EXCEL_FILE, "rb") as f:
        st.download_button("Descargar Excel", data=f.read(), file_name="facturas.xlsx", type="primary")