"""
Módulo para subir archivos a SharePoint via Microsoft Graph API.
La autenticación se maneja externamente (token pasado como parámetro).
Requiere: pip install msal requests
"""

import os
import json
import uuid
import tempfile
import requests
import msal

SITE_HOSTNAME = "catodica.sharepoint.com"
SITE_PATH = "/sites/ProyectoCips-PowerBI"
LIBRARY_NAME = "Resultados_CIPS"
SCOPES = ["https://graph.microsoft.com/Files.ReadWrite.All"]


def iniciar_auth_code_flow(tenant_id, client_id, client_secret, redirect_uri):
    """Genera la URL de login de Microsoft y guarda el flow en disco."""
    app_auth = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    flow = app_auth.initiate_auth_code_flow(scopes=SCOPES, redirect_uri=redirect_uri)
    if "auth_uri" not in flow:
        raise Exception(f"Error iniciando login: {flow}")
    state = flow.get("state", str(uuid.uuid4()))
    flow_path = os.path.join(tempfile.gettempdir(), f"msal_flow_{state}.json")
    with open(flow_path, "w") as f:
        json.dump(flow, f)
    return flow["auth_uri"]


def completar_auth_code_flow(query_params, tenant_id, client_id, client_secret):
    """Completa el login con los parámetros que Microsoft devuelve en la URL."""
    state = query_params.get("state", "")
    flow_path = os.path.join(tempfile.gettempdir(), f"msal_flow_{state}.json")
    if not os.path.exists(flow_path):
        raise Exception("Sesión de login expirada. Haz clic en Login de nuevo.")
    with open(flow_path) as f:
        flow = json.load(f)
    try:
        os.unlink(flow_path)
    except Exception:
        pass
    app_auth = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    result = app_auth.acquire_token_by_auth_code_flow(flow, query_params)
    if "access_token" not in result:
        raise Exception(f"Error autenticando: {result.get('error_description', result)}")
    return result["access_token"]


def login_interactivo(tenant_id, client_id):
    """Abre el navegador directamente para login. Retorna el access token."""
    app_auth = msal.PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    # Reutilizar cuenta existente si hay
    cuentas = app_auth.get_accounts()
    if cuentas:
        result = app_auth.acquire_token_silent(SCOPES, account=cuentas[0])
        if result and "access_token" in result:
            return result["access_token"]

    result = app_auth.acquire_token_interactive(scopes=SCOPES)
    if "access_token" not in result:
        raise Exception(f"Error autenticando: {result.get('error_description', result)}")
    return result["access_token"]


def subir_a_sharepoint(ruta_archivo, token, subcarpeta=None):
    """
    Sube un archivo a la librería Resultados_CIPS en SharePoint.

    ruta_archivo: ruta local del archivo
    token:        access token de Microsoft Graph
    subcarpeta:   carpeta dentro de la librería (ej: "2025/04"), opcional

    Retorna la URL web del archivo subido.
    """
    headers = {"Authorization": f"Bearer {token}"}

    # Obtener ID del sitio
    site_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{SITE_HOSTNAME}:{SITE_PATH}",
        headers=headers
    )
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    # Obtener ID del drive (librería de documentos)
    drives_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        headers=headers
    )
    drives_resp.raise_for_status()
    drives = drives_resp.json().get("value", [])

    drive_id = next((d["id"] for d in drives if d["name"] == LIBRARY_NAME), None)
    if not drive_id:
        nombres = [d["name"] for d in drives]
        raise Exception(
            f"No se encontró la librería '{LIBRARY_NAME}' en SharePoint.\n"
            f"Librerías disponibles: {nombres}"
        )

    # Construir ruta de destino
    nombre_archivo = os.path.basename(ruta_archivo)
    ruta_destino = f"{subcarpeta}/{nombre_archivo}" if subcarpeta else nombre_archivo

    # Subir archivo (simple para <4 MB, upload session para archivos grandes)
    file_size = os.path.getsize(ruta_archivo)

    if file_size < 4 * 1024 * 1024:
        upload_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
            f"/root:/{ruta_destino}:/content"
        )
        with open(ruta_archivo, "rb") as f:
            resp = requests.put(
                upload_url,
                headers={**headers, "Content-Type": "application/octet-stream"},
                data=f.read()
            )
    else:
        session_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
            f"/root:/{ruta_destino}:/createUploadSession"
        )
        session_resp = requests.post(session_url, headers=headers, json={})
        session_resp.raise_for_status()
        upload_url = session_resp.json()["uploadUrl"]

        chunk_size = 10 * 1024 * 1024
        resp = None
        with open(ruta_archivo, "rb") as f:
            offset = 0
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                end = offset + len(chunk) - 1
                resp = requests.put(
                    upload_url,
                    headers={
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {offset}-{end}/{file_size}"
                    },
                    data=chunk
                )
                offset += len(chunk)

    if resp.status_code not in [200, 201]:
        raise Exception(f"Error HTTP {resp.status_code}: {resp.text}")

    return resp.json().get("webUrl", "")
