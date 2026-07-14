#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
refresh_from_teams.py
=====================
Motor de auto-actualización del dashboard PMO.

Se conecta a SharePoint/Teams con Microsoft Graph (autenticación de aplicación,
client-credentials), busca dentro de la carpeta de **Seguimiento Semanal** el
archivo Excel cuya **fecha en el nombre sea la más cercana a HOY**, lo descarga
y regenera `dashboard.html` con `build_dashboard.py`.

Carpeta objetivo:
    TECNICO-BOULEVARSUR › PLANIFICACION › 01. PLANNER › FASE 2 › 02. Seguimiento Semanal

Variables de entorno requeridas (se configuran como GitHub Secrets):
    AZURE_TENANT_ID       - ID del inquilino (tenant) de Azure AD
    AZURE_CLIENT_ID       - ID de la aplicación registrada
    AZURE_CLIENT_SECRET   - secreto de cliente de la aplicación
Opcionales (ya traen el valor correcto por defecto para esta carpeta):
    GRAPH_DRIVE_ID        - ID de la biblioteca de documentos
    GRAPH_FOLDER_ID       - ID de la carpeta "02. Seguimiento Semanal"

La app necesita el permiso de aplicación **Sites.Read.All** (o Files.Read.All)
con consentimiento de administrador. Es de SOLO LECTURA.

Uso:
    python3 refresh_from_teams.py
"""
import os, re, sys, datetime, urllib.request, urllib.parse, json

import build_dashboard  # reutiliza extract_from_xlsx / enrich / render

# ---- Identificadores de la carpeta de Teams (verificados vía Graph) ---------
DRIVE_ID  = os.environ.get("GRAPH_DRIVE_ID",
    "b!wd7ZI735FkuclbJhZDMyeIlTSAqOzsZNm_G2WbyYKEpaEDioIdTHSpf-IrdwUQba")
FOLDER_ID = os.environ.get("GRAPH_FOLDER_ID",
    "01CKHUMJ4GKZOM42LYRRAZPATSOX7KJX6E")

GRAPH = "https://graph.microsoft.com/v1.0"


def _http_json(url, data=None, headers=None):
    req = urllib.request.Request(url, data=data, headers=headers or {})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode())


def get_token():
    tenant = os.environ["AZURE_TENANT_ID"]
    body = urllib.parse.urlencode({
        "client_id":     os.environ["AZURE_CLIENT_ID"],
        "client_secret": os.environ["AZURE_CLIENT_SECRET"],
        "scope":         "https://graph.microsoft.com/.default",
        "grant_type":    "client_credentials",
    }).encode()
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    tok = _http_json(url, data=body,
                     headers={"Content-Type": "application/x-www-form-urlencoded"})
    return tok["access_token"]


def parse_date_from_name(name):
    """Extrae la fecha del nombre del archivo. Soporta '13 07 2026' y '2026-06-18'."""
    m = re.search(r"(20\d{2})[ _\-.](\d{1,2})[ _\-.](\d{1,2})", name)   # YYYY MM DD
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.search(r"(\d{1,2})[ _\-.](\d{1,2})[ _\-.](20\d{2})", name)  # DD MM YYYY
        if not m:
            return None
        d, mo, y = map(int, m.groups())
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return None


def list_children(token):
    url = f"{GRAPH}/drives/{DRIVE_ID}/items/{FOLDER_ID}/children?$top=200"
    items, out = [], []
    data = _http_json(url, headers={"Authorization": f"Bearer {token}"})
    items.extend(data.get("value", []))
    while data.get("@odata.nextLink"):
        data = _http_json(data["@odata.nextLink"],
                          headers={"Authorization": f"Bearer {token}"})
        items.extend(data.get("value", []))
    for it in items:
        if "file" not in it:            # ignorar subcarpetas
            continue
        out.append(it)
    return out


def pick_closest_xlsx(items, today=None):
    """Elige el .xlsx cuya fecha en el nombre sea la más cercana a hoy."""
    today = today or datetime.date.today()
    cands = []
    for it in items:
        name = it.get("name", "")
        if not name.lower().endswith(".xlsx"):
            continue
        d = parse_date_from_name(name)
        if d is None:
            # sin fecha en el nombre -> usar fecha de modificación como respaldo
            lm = it.get("lastModifiedDateTime", "")[:10]
            try:
                d = datetime.date.fromisoformat(lm)
            except ValueError:
                continue
        cands.append((abs((d - today).days), -d.toordinal(), it, d))
    if not cands:
        return None, None
    cands.sort(key=lambda t: (t[0], t[1]))   # más cercano; desempate: fecha más reciente
    _, _, it, d = cands[0]
    return it, d


def download(token, item, dest):
    dl = item.get("@microsoft.graph.downloadUrl")
    if not dl:
        meta = _http_json(f"{GRAPH}/drives/{DRIVE_ID}/items/{item['id']}",
                          headers={"Authorization": f"Bearer {token}"})
        dl = meta.get("@microsoft.graph.downloadUrl")
    req = urllib.request.Request(dl)
    with urllib.request.urlopen(req, timeout=120) as r, open(dest, "wb") as fh:
        fh.write(r.read())
    return dest


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    try:
        token = get_token()
    except KeyError as e:
        sys.exit(f"[error] falta la variable de entorno {e}. "
                 f"Configura AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET.")

    items = list_children(token)
    print(f"[ok] {len(items)} archivos en la carpeta de Seguimiento Semanal")

    item, fdate = pick_closest_xlsx(items)
    if not item:
        sys.exit("[error] no se encontró ningún .xlsx con fecha en la carpeta.")
    print(f"[ok] archivo elegido: {item['name']}  (fecha {fdate}, "
          f"a {abs((fdate - datetime.date.today()).days)} días de hoy)")

    xlsx = os.path.join(here, "_ultimo_seguimiento.xlsx")
    download(token, item, xlsx)
    print(f"[ok] descargado -> {xlsx} ({os.path.getsize(xlsx)} bytes)")

    data = build_dashboard.extract_from_xlsx(xlsx)
    data = build_dashboard.enrich(data)
    html = build_dashboard.render(data)
    out = os.path.join(here, "dashboard.html")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(html)
    p = data["proyecto"]
    print(f"[ok] dashboard.html actualizado · corte {data['corte']} · "
          f"{p.get('tareas','?')} tareas · avance {p.get('avance')}")
    try:
        os.remove(xlsx)
    except OSError:
        pass


if __name__ == "__main__":
    main()
