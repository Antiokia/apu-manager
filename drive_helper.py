"""
drive_helper.py
---------------
Sube un archivo a una carpeta específica de Google Drive
usando una Service Account configurada en st.secrets.

Configuración en Streamlit Cloud → Secrets:
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\n..."
client_email = "..."
client_id = "..."
...

[drive]
folder_id = "ID_DE_TU_CARPETA_EN_DRIVE"
"""

import io
import json
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


SCOPES = ["https://www.googleapis.com/auth/drive.file"]


def _get_service():
    """Construye el cliente de Drive desde st.secrets."""
    info = dict(st.secrets["gcp_service_account"])
    # Streamlit escapa los \n en la clave privada — los restauramos
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def subir_a_drive(file_bytes: bytes, nombre_archivo: str) -> str:
    """
    Sube file_bytes a la carpeta configurada en st.secrets[drive][folder_id].
    Devuelve el link público del archivo en Drive.
    """
    service   = _get_service()
    folder_id = st.secrets["drive"]["folder_id"]

    # Detectar MIME type por extensión
    ext = nombre_archivo.rsplit(".", 1)[-1].lower()
    mime_map = {
        "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "xls":  "application/vnd.ms-excel",
    }
    mime_type = mime_map.get(ext, "application/octet-stream")

    file_metadata = {
        "name":    nombre_archivo,
        "parents": [folder_id],
    }
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)

    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, webViewLink",
    ).execute()

    return uploaded.get("webViewLink", "")


def drive_configurado() -> bool:
    """True si existen los secrets necesarios."""
    try:
        _ = st.secrets["gcp_service_account"]
        _ = st.secrets["drive"]["folder_id"]
        return True
    except Exception:
        return False
