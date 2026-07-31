#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
minutas_sources.py
==================
Conectores a las tres fuentes de minutas/reuniones. Cada conector devuelve una
lista de reuniones NORMALIZADAS con este esquema común:

    {
      "source":       "teams" | "meetgeek" | "otter",
      "meeting_id":   str,
      "title":        str,
      "date":         "YYYY-MM-DD" | "",
      "attendees":    [str, ...],
      "summary":      str,              # resumen/notas si la fuente lo da
      "text":         str,              # transcripción o notas en texto plano
      "action_items": [ {"text":str, "owner":str, "due":str}, ... ],  # nativos
      "url":          str,
    }

Diseño defensivo: si a una fuente le faltan credenciales, su `fetch_*` devuelve
[] y anota el motivo en `skipped`; NUNCA lanza excepción que tumbe el proceso.
Así el orquestador corre con las fuentes que sí estén configuradas.

Credenciales (variables de entorno / GitHub Secrets):
  Teams (Graph):  AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET
                  + GRAPH_MINUTAS_FOLDER_ID (carpeta de SharePoint con las minutas)
                  + GRAPH_MINUTAS_DRIVE_ID  (opcional; por defecto el mismo drive del PMO)
  MeetGeek:       MEETGEEK_API_KEY
  Otter.ai:       OTTER_EMAIL / OTTER_PASSWORD
"""
import os, re, json, datetime, urllib.request, urllib.parse, urllib.error, base64

import graph_common as gc

skipped = []   # motivos por los que se saltó alguna fuente (para el log/HTML)


# --------------------------------------------------------------------------- #
#  Utilidades                                                                   #
# --------------------------------------------------------------------------- #
def _iso_date(s):
    if not s:
        return ""
    s = str(s)[:10]
    try:
        datetime.date.fromisoformat(s)
        return s
    except ValueError:
        return ""


def _get(url, headers=None, timeout=60):
    return gc.http_json(url, headers=headers, timeout=timeout)


def _text_from_bytes(name, raw):
    """Extrae texto plano de .txt/.vtt/.md/.docx/.pdf (best-effort, sin libs pesadas)."""
    low = name.lower()
    if low.endswith((".txt", ".md", ".vtt", ".csv")):
        t = raw.decode("utf-8", "replace")
        if low.endswith(".vtt"):   # limpia marcas de tiempo de WebVTT
            t = "\n".join(l for l in t.splitlines()
                          if "-->" not in l and not l.strip().isdigit()
                          and l.strip().upper() != "WEBVTT")
        return t
    if low.endswith(".docx"):
        try:
            import zipfile, io, xml.etree.ElementTree as ET
            with zipfile.ZipFile(io.BytesIO(raw)) as z:
                xml = z.read("word/document.xml")
            ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            root = ET.fromstring(xml)
            paras = []
            for p in root.iter(ns + "p"):
                txt = "".join(t.text or "" for t in p.iter(ns + "t"))
                if txt.strip():
                    paras.append(txt)
            return "\n".join(paras)
        except Exception:
            return ""
    return ""   # .pdf y otros: se ignoran silenciosamente (evitar dependencias)


# --------------------------------------------------------------------------- #
#  1) Teams / SharePoint (Microsoft Graph, solo lectura)                        #
# --------------------------------------------------------------------------- #
def fetch_teams(limit=40):
    """Lee minutas de una carpeta de SharePoint/Teams (Word/txt/vtt/md).

    La carpeta la define GRAPH_MINUTAS_FOLDER_ID. Cada archivo se toma como una
    'reunión' cuyo texto son sus notas/transcripción. Ideal para pegar ahí lo
    que Teams/MeetGeek/Otter exporten, o las notas que tome el equipo.
    """
    if not gc.have_graph_creds():
        skipped.append("Teams: faltan AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET")
        return []
    folder = gc.clean("GRAPH_MINUTAS_FOLDER_ID")
    if not folder:
        skipped.append("Teams: falta GRAPH_MINUTAS_FOLDER_ID (carpeta de minutas en SharePoint)")
        return []
    drive = gc.clean("GRAPH_MINUTAS_DRIVE_ID") or gc.clean("GRAPH_DRIVE_ID") or \
        "b!wd7ZI735FkuclbJhZDMyeIlTSAqOzsZNm_G2WbyYKEpaEDioIdTHSpf-IrdwUQba"
    try:
        token = gc.get_token()
    except SystemExit as e:
        skipped.append(f"Teams: {e}")
        return []
    h = {"Authorization": f"Bearer {token}"}
    url = f"{gc.GRAPH}/drives/{drive}/items/{folder}/children?$top=200&$orderby=lastModifiedDateTime desc"
    items = []
    try:
        data = _get(url, headers=h)
        items.extend(data.get("value", []))
        while data.get("@odata.nextLink") and len(items) < 400:
            data = _get(data["@odata.nextLink"], headers=h)
            items.extend(data.get("value", []))
    except urllib.error.HTTPError as e:
        skipped.append(f"Teams: Graph respondió HTTP {e.code} al leer la carpeta de minutas")
        return []

    out = []
    for it in items:
        if "file" not in it:
            continue
        name = it.get("name", "")
        if not name.lower().endswith((".txt", ".md", ".vtt", ".docx", ".csv")):
            continue
        dl = it.get("@microsoft.graph.downloadUrl")
        text = ""
        if dl:
            try:
                with urllib.request.urlopen(dl, timeout=90) as r:
                    text = _text_from_bytes(name, r.read())
            except Exception:
                text = ""
        d = _parse_date_from_name(name) or _iso_date(it.get("lastModifiedDateTime", ""))
        out.append({
            "source": "teams", "meeting_id": it.get("id", name),
            "title": re.sub(r"\.[^.]+$", "", name), "date": d,
            "attendees": [], "summary": "", "text": text,
            "action_items": [], "url": it.get("webUrl", ""),
        })
        if len(out) >= limit:
            break
    return out


def _parse_date_from_name(name):
    m = re.search(r"(20\d{2})[ _\-.](\d{1,2})[ _\-.](\d{1,2})", name)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.search(r"(\d{1,2})[ _\-.](\d{1,2})[ _\-.](20\d{2})", name)
        if not m:
            return ""
        d, mo, y = map(int, m.groups())
    try:
        return datetime.date(y, mo, d).isoformat()
    except ValueError:
        return ""


# --------------------------------------------------------------------------- #
#  2) MeetGeek (API REST v1, Bearer)                                            #
# --------------------------------------------------------------------------- #
MEETGEEK = "https://api.meetgeek.ai/v1"


def fetch_meetgeek(limit=25):
    key = gc.clean("MEETGEEK_API_KEY")
    if not key:
        skipped.append("MeetGeek: falta MEETGEEK_API_KEY")
        return []
    h = {"Authorization": f"Bearer {key}", "Accept": "application/json"}
    out = []
    try:
        page = _get(f"{MEETGEEK}/meetings?count={limit}", headers=h)
    except urllib.error.HTTPError as e:
        skipped.append(f"MeetGeek: API respondió HTTP {e.code} (revisa MEETGEEK_API_KEY)")
        return []
    meetings = page.get("values", page.get("data", page if isinstance(page, list) else []))
    for m in meetings[:limit]:
        mid = m.get("meeting_id") or m.get("id")
        if not mid:
            continue
        title = m.get("title") or m.get("name") or "Reunión MeetGeek"
        date = _iso_date(m.get("timestamp_start_utc") or m.get("start_time") or m.get("date"))
        summary, ai, text = "", [], ""
        try:
            s = _get(f"{MEETGEEK}/meetings/{mid}/summary", headers=h)
            summary = s.get("summary") or s.get("text") or ""
        except Exception:
            pass
        try:
            hl = _get(f"{MEETGEEK}/meetings/{mid}/highlights", headers=h)
            for item in hl.get("values", hl.get("data", [])):
                txt = item.get("text") or item.get("highlight") or ""
                cat = (item.get("category") or item.get("type") or "").lower()
                if txt and ("action" in cat or "task" in cat or "todo" in cat):
                    ai.append({"text": txt, "owner": item.get("owner", ""), "due": ""})
                elif txt:
                    text += txt + "\n"
        except Exception:
            pass
        out.append({
            "source": "meetgeek", "meeting_id": str(mid), "title": title, "date": date,
            "attendees": m.get("participants", []) or [], "summary": summary,
            "text": text or summary, "action_items": ai,
            "url": m.get("url", f"https://app.meetgeek.ai/meeting/{mid}"),
        })
    return out


# --------------------------------------------------------------------------- #
#  3) Otter.ai (API no oficial usada por la web, login + speeches)              #
# --------------------------------------------------------------------------- #
OTTER = "https://otter.ai/forward/api/v1"


def _otter_login(email, password):
    userpass = base64.b64encode(f"{email}:{password}".encode()).decode()
    url = (f"{OTTER}/login?username=" + urllib.parse.quote(email))
    req = urllib.request.Request(url, headers={"Authorization": f"Basic {userpass}"})
    cj = urllib.request.HTTPCookieProcessor()
    opener = urllib.request.build_opener(cj)
    with opener.open(req, timeout=60) as r:
        r.read()
    return opener, cj


def fetch_otter(limit=25):
    email = gc.clean("OTTER_EMAIL")
    password = gc.clean("OTTER_PASSWORD")
    if not (email and password):
        skipped.append("Otter.ai: faltan OTTER_EMAIL / OTTER_PASSWORD "
                       "(en esta sesión interactiva se puede leer por el conector MCP de Otter)")
        return []
    try:
        opener, _ = _otter_login(email, password)
        url = f"{OTTER}/speeches?page_size={limit}"
        with opener.open(url, timeout=60) as r:
            data = json.loads(r.read().decode())
    except Exception as e:
        skipped.append(f"Otter.ai: no se pudo autenticar/leer ({type(e).__name__})")
        return []
    out = []
    for sp in (data.get("speeches") or [])[:limit]:
        sid = sp.get("otid") or sp.get("id")
        title = sp.get("title") or "Reunión Otter"
        ts = sp.get("start_time") or sp.get("created_at")
        try:
            date = datetime.datetime.utcfromtimestamp(int(ts)).date().isoformat() if ts else ""
        except (ValueError, TypeError):
            date = _iso_date(ts)
        text = sp.get("summary") or ""
        try:
            with opener.open(f"{OTTER}/speech?otid={sid}", timeout=60) as r:
                det = json.loads(r.read().decode())
            trs = det.get("speech", {}).get("transcripts", [])
            text = (text + "\n" + "\n".join(t.get("transcript", "") for t in trs)).strip()
        except Exception:
            pass
        out.append({
            "source": "otter", "meeting_id": str(sid), "title": title, "date": date,
            "attendees": [], "summary": sp.get("summary", ""), "text": text,
            "action_items": [], "url": f"https://otter.ai/u/{sid}",
        })
    return out


# --------------------------------------------------------------------------- #
def fetch_all(limit_each=30):
    """Junta las reuniones de todas las fuentes configuradas."""
    skipped.clear()
    meetings = []
    meetings += fetch_teams(limit_each)
    meetings += fetch_meetgeek(limit_each)
    meetings += fetch_otter(limit_each)
    meetings.sort(key=lambda m: m.get("date", ""), reverse=True)
    return meetings


if __name__ == "__main__":
    ms = fetch_all()
    print(f"[ok] {len(ms)} reuniones de fuentes configuradas")
    for s in skipped:
        print("  · (saltada) " + s)
    for m in ms[:10]:
        print(f"  - [{m['source']}] {m['date']} · {m['title']} "
              f"({len(m['action_items'])} acciones nativas, {len(m['text'])} chars)")
