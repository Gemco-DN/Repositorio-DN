# -*- coding: utf-8 -*-
"""
Descarga de archivos Excel desde SharePoint vía Microsoft Graph API.

Mismo patrón (MSAL, Client Credentials Flow) usado en Panel Licitaciones
(actualizar_panel_v2.py). Reutiliza las mismas credenciales de Azure AD
(SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET) — solo cambia el link de
"Compartir" de cada archivo.

Variables de entorno requeridas:
- SP_TENANT_ID     : Tenant ID de Entra ID
- SP_CLIENT_ID     : Client ID del App Registration
- SP_CLIENT_SECRET : Client Secret del App Registration
"""

import base64
import os
from io import BytesIO

import msal
import requests

SP_TENANT_ID = os.environ["SP_TENANT_ID"]
SP_CLIENT_ID = os.environ["SP_CLIENT_ID"]
SP_CLIENT_SECRET = os.environ["SP_CLIENT_SECRET"]

GRAPH_AUTHORITY = f"https://login.microsoftonline.com/{SP_TENANT_ID}"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

_token_cache = {}


def _obtener_token_graph():
    """Autentica contra Entra ID vía MSAL y devuelve un access token para Graph API.
    Se cachea en memoria del proceso para no re-autenticar por cada archivo."""
    if "token" in _token_cache:
        return _token_cache["token"]
    app = msal.ConfidentialClientApplication(
        SP_CLIENT_ID, authority=GRAPH_AUTHORITY, client_credential=SP_CLIENT_SECRET
    )
    resultado = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in resultado:
        raise RuntimeError(
            f"No se pudo obtener token de Graph API: "
            f"{resultado.get('error')} - {resultado.get('error_description')}"
        )
    _token_cache["token"] = resultado["access_token"]
    return resultado["access_token"]


def _codificar_url_para_graph(url):
    """Codifica una URL de SharePoint en el formato 'sharing token' que espera /shares.
    Ref: https://learn.microsoft.com/graph/api/shares-get"""
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return "u!" + b64


def descargar_excel(share_url, nombre_archivo="archivo"):
    """Descarga un Excel desde SharePoint vía Graph API y lo devuelve en memoria (BytesIO)."""
    token = _obtener_token_graph()
    headers = {"Authorization": f"Bearer {token}"}
    share_id = _codificar_url_para_graph(share_url)

    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem",
        headers=headers, timeout=30,
    )
    resp.raise_for_status()
    drive_item = resp.json()

    download_url = drive_item.get("@microsoft.graph.downloadUrl")
    if download_url:
        contenido = requests.get(download_url, timeout=60)
    else:
        drive_id = drive_item["parentReference"]["driveId"]
        item_id = drive_item["id"]
        contenido = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
            headers=headers, timeout=60,
        )
    contenido.raise_for_status()
    print(f"[OK] Descargado desde SharePoint: {nombre_archivo}")
    return BytesIO(contenido.content)
