#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
graph_common.py
===============
Utilidades compartidas de Microsoft Graph para los módulos del proyecto
(dashboard PMO y el nuevo módulo de Minutas & Acciones).

Centraliza:
  * obtención del token de aplicación (client-credentials) con validación y
    diagnóstico de errores comunes de Azure (mismo comportamiento que ya tenía
    `refresh_from_teams.py`);
  * pequeños helpers HTTP JSON.

Variables de entorno:
  AZURE_TENANT_ID       - inquilino (GUID o dominio .onmicrosoft.com)
  AZURE_CLIENT_ID       - id de aplicación (GUID)
  AZURE_CLIENT_SECRET   - VALOR del secreto (no el Id. del secreto)

Para SOLO LECTURA basta el permiso de aplicación Sites.Read.All / Files.Read.All.
Para ESCRIBIR tareas en Planner (opcional) la app necesita, además,
Group.ReadWrite.All (o Tasks.ReadWrite) con consentimiento de administrador.
El token es el mismo (scope .default): los permisos se otorgan en Azure, no aquí.
"""
import os, re, sys, json, urllib.request, urllib.parse, urllib.error

GRAPH = "https://graph.microsoft.com/v1.0"


def clean(name):
    """Lee un secreto y le quita espacios/comillas/saltos accidentales."""
    return (os.environ.get(name, "") or "").strip().strip('"').strip("'").strip()


def http_json(url, data=None, headers=None, method=None, timeout=60):
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        body = r.read().decode()
        return json.loads(body) if body else {}


def get_token():
    """Token de aplicación de Microsoft Graph (client-credentials)."""
    tenant = clean("AZURE_TENANT_ID")
    client_id = clean("AZURE_CLIENT_ID")
    client_secret = clean("AZURE_CLIENT_SECRET")
    guid = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
                      r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
    domain = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.-]*\.[A-Za-z]{2,}$")

    def diag(v):
        return (f"[diagnóstico sin exponer el valor: {len(v)} caracteres, "
                f"con espacio={'SÍ' if any(c.isspace() for c in v) else 'no'}, "
                f"empieza con='{(v[:1] or '∅')}']")

    if not (guid.match(tenant) or domain.match(tenant)):
        sys.exit("[error] AZURE_TENANT_ID inválido. Debe ser el Id. de directorio "
                 "(GUID) o el dominio .onmicrosoft.com. " + diag(tenant))
    if not guid.match(client_id):
        sys.exit("[error] AZURE_CLIENT_ID inválido: debe ser el Id. de aplicación "
                 "(GUID). " + diag(client_id))
    if guid.match(client_secret):
        sys.exit("[error] AZURE_CLIENT_SECRET parece ser el 'Id. del secreto' "
                 "(GUID), NO el 'Valor'. Copia la columna 'Valor'. " + diag(client_secret))

    body = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode()
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    try:
        tok = http_json(url, data=body,
                        headers={"Content-Type": "application/x-www-form-urlencoded"})
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")
        try:
            j = json.loads(detail)
            code = j.get("error", "")
            desc = (j.get("error_description", "") or "").split("\r\n")[0]
        except Exception:
            code, desc = "", detail[:300]
        hint = ""
        if "7000215" in detail:
            hint = "→ El CLIENT SECRET no coincide (¿pegaste el 'Id.' en vez del 'Valor'?)."
        elif "7000222" in detail:
            hint = "→ El client secret EXPIRÓ. Crea uno nuevo en Azure."
        elif "700016" in detail:
            hint = "→ La app no existe en este tenant (revisa CLIENT_ID/TENANT_ID)."
        sys.exit(f"[error] Azure rechazó la autenticación (HTTP {e.code}). {code}: {desc} {hint}")
    return tok["access_token"]


def have_graph_creds():
    return bool(clean("AZURE_TENANT_ID") and clean("AZURE_CLIENT_ID")
                and clean("AZURE_CLIENT_SECRET"))
