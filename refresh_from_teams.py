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
import os, re, sys, datetime, urllib.request, urllib.parse, urllib.error, json

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


def _clean(name):
    """lee un secreto y le quita espacios/comillas/saltos de línea accidentales."""
    return (os.environ.get(name, "") or "").strip().strip('"').strip("'").strip()

def get_token():
    tenant = _clean("AZURE_TENANT_ID")
    client_id = _clean("AZURE_CLIENT_ID")
    client_secret = _clean("AZURE_CLIENT_SECRET")
    # Validaciones con mensajes claros. El tenant acepta GUID o dominio
    # (p.ej. periferiaurbana.onmicrosoft.com); el client debe ser GUID.
    guid = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
                      r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
    domain = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.-]*\.[A-Za-z]{2,}$")
    def diag(v):
        return (f"[diagnóstico sin exponer el valor: {len(v)} caracteres, "
                f"con espacio={'SÍ' if any(c.isspace() for c in v) else 'no'}, "
                f"con punto={'sí' if '.' in v else 'no'}, "
                f"empieza con='{(v[:1] or '∅')}']")
    if not (guid.match(tenant) or domain.match(tenant)):
        sys.exit("[error] AZURE_TENANT_ID inválido. Debe ser el 'Id. de directorio (inquilino)' "
                 "(GUID como a1b2c3d4-e5f6-7890-abcd-1234567890ab) o el dominio "
                 "(p.ej. periferiaurbana.onmicrosoft.com), sin espacios. " + diag(tenant))
    if not guid.match(client_id):
        sys.exit("[error] AZURE_CLIENT_ID inválido: debe ser el 'Id. de aplicación (cliente)' "
                 "(GUID). " + diag(client_id))
    body = urllib.parse.urlencode({
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
        "grant_type":    "client_credentials",
    }).encode()
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    try:
        tok = _http_json(url, data=body,
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
            hint = ("→ El CLIENT SECRET es incorrecto. En Azure, "
                    "'Certificados y secretos', copia el campo 'Valor' del secreto "
                    "(NO el 'Id. del secreto') y actualízalo en AZURE_CLIENT_SECRET.")
        elif "7000222" in detail:
            hint = "→ El client secret EXPIRÓ. Crea uno nuevo en Azure y actualiza AZURE_CLIENT_SECRET."
        elif "700016" in detail:
            hint = "→ La app no existe en este tenant: revisa AZURE_CLIENT_ID y AZURE_TENANT_ID."
        elif "900023" in detail or "90002" in detail:
            hint = "→ El tenant no es válido: revisa AZURE_TENANT_ID."
        sys.exit(f"[error] Azure rechazó la autenticación (HTTP {e.code}). "
                 f"{code}: {desc} {hint}")
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
