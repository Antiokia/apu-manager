"""
drive_helper.py  —  OAuth 2.0 (funciona con Google Drive personal)
-------------------------------------------------------------------
Flujo:
  1. Primera vez: genera un URL de autorización → el admin lo abre,
     autoriza, copia el código de vuelta a Secrets.
  2. De ahí en adelante: usa el refresh_token guardado en Secrets.

Secrets necesarios en Streamlit Cloud:
---------------------------------------
[oauth]
client_id     = "xxxx.apps.googleusercontent.com"
client_secret = "GOCSPX-..."
refresh_token = ""          # vacío la primera vez

[drive]
folder_id = "ID_DE_TU_CARPETA"
"""

import io
import json
import streamlit as st
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google_auth_oauthlib.flow import Flow

SCOPES = ["https://www.googleapis.com/auth/drive.file"]
REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob"   # modo "copiar código"


# ─────────────────────────────────────────────
# HELPERS DE SECRETS
# ─────────────────────────────────────────────

def _oauth_secrets():
    return st.secrets["oauth"]

def _folder_id():
    return st.secrets["drive"]["folder_id"]

def drive_configurado() -> bool:
    try:
        s = _oauth_secrets()
        _ = s["client_id"]
        _ = s["client_secret"]
        _ = _folder_id()
        return True
    except Exception:
        return False

def _tiene_refresh_token() -> bool:
    try:
        rt = _oauth_secrets().get("refresh_token", "")
        return bool(rt and str(rt).strip())
    except Exception:
        return False


# ─────────────────────────────────────────────
# OBTENER CREDENCIALES
# ─────────────────────────────────────────────

def _get_credentials() -> Credentials:
    s = _oauth_secrets()
    creds = Credentials(
        token         = None,
        refresh_token = str(s["refresh_token"]).strip(),
        token_uri     = "https://oauth2.googleapis.com/token",
        client_id     = str(s["client_id"]).strip(),
        client_secret = str(s["client_secret"]).strip(),
        scopes        = SCOPES,
    )
    # Refrescar el access token
    creds.refresh(Request())
    return creds


def _get_service():
    creds = _get_credentials()
    return build("drive", "v3", credentials=creds, cache_discovery=False)


# ─────────────────────────────────────────────
# GENERAR URL DE AUTORIZACIÓN (primera vez)
# ─────────────────────────────────────────────

def generar_url_autorizacion() -> str:
    s = _oauth_secrets()
    client_config = {
        "installed": {
            "client_id":     str(s["client_id"]).strip(),
            "client_secret": str(s["client_secret"]).strip(),
            "auth_uri":      "https://accounts.google.com/o/oauth2/auth",
            "token_uri":     "https://oauth2.googleapis.com/token",
            "redirect_uris": [REDIRECT_URI],
        }
    }
    flow = Flow.from_client_config(
        client_config,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        prompt="consent",
    )
    return auth_url


def obtener_refresh_token(codigo: str) -> str:
    """Intercambia el código de autorización por un refresh_token."""
    s = _oauth_secrets()
    client_config = {
        "installed": {
            "client_id":     str(s["client_id"]).strip(),
            "client_secret": str(s["client_secret"]).strip(),
            "auth_uri":      "https://accounts.google.com/o/oauth2/auth",
            "token_uri":     "https://oauth2.googleapis.com/token",
            "redirect_uris": [REDIRECT_URI],
        }
    }
    flow = Flow.from_client_config(
        client_config,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    flow.fetch_token(code=codigo.strip())
    return flow.credentials.refresh_token


# ─────────────────────────────────────────────
# SUBIR ARCHIVO
# ─────────────────────────────────────────────

def subir_a_drive(file_bytes: bytes, nombre_archivo: str) -> str:
    service   = _get_service()
    folder_id = _folder_id()

    ext = nombre_archivo.rsplit(".", 1)[-1].lower()
    mime_map = {
        "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "xls":  "application/vnd.ms-excel",
    }
    mime_type = mime_map.get(ext, "application/octet-stream")

    file_metadata = {"name": nombre_archivo, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)

    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, webViewLink",
    ).execute()

    return uploaded.get("webViewLink", "")
