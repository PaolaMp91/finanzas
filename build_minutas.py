#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_minutas.py
================
Convierte minutas de reuniones (Teams / MeetGeek / Otter.ai, ya normalizadas por
`minutas_sources.py`) en un tablero HTML autocontenido `minutas.html` con tres
paneles:

  1) ASIGNACIÓN DE TAREAS  — quién hace qué, para cuándo, de qué reunión salió.
  2) DECISIONES DE SOCIOS  — acuerdos/decisiones tomadas por los socios.
  3) HISTORIAL FINANCIERO  — cierres contables y escenarios financieros mencionados.

Uso:
    python3 build_minutas.py                # vista previa con datos de demo (SEED)
    python3 build_minutas.py --stdin        # lee reuniones (JSON) por stdin
El orquestador `refresh_minutas.py` llama a extract()/enrich()/render() con las
reuniones reales.
"""
import os, re, sys, json, datetime, unicodedata

# --------------------------------------------------------------------------- #
#  Configuración del equipo / socios (ajústalo a tu organización)              #
# --------------------------------------------------------------------------- #
# Socios: sus decisiones alimentan el panel "Decisiones de socios".
SOCIOS = ["Paola", "Socio 1", "Socio 2", "Junta"]

# Normalización de nombres de responsables (alias -> nombre canónico).
RESP_ALIAS = {
    "pao": "Paola", "paola": "Paola",
    "ing": "Ingeniería", "ingenieria": "Ingeniería",
    "conta": "Contabilidad", "contabilidad": "Contabilidad", "finanzas": "Contabilidad",
    "pmo": "PMO", "planificacion": "PMO", "planificación": "PMO",
    "ventas": "Ventas", "comercial": "Ventas",
}

# Palabras clave (sin acentos, minúsculas) para clasificar líneas de la minuta.
KW_TAREA = ["tarea", "responsable", "encargad", "queda de", "se compromete",
            "hara", "enviar", "entregar", "preparar", "revisar", "coordinar",
            "seguimiento", "pendiente", "to do", "todo", "action", "asignad"]
KW_DECISION = ["se decide", "se decidio", "decidimos", "acuerda", "acordamos",
               "acordado", "se aprueba", "aprobado", "aprobamos", "resolucion",
               "resuelve", "queda aprobado", "los socios", "la junta", "acuerdo:"]
KW_FIN = ["cierre contable", "cierre de mes", "cierre mensual", "cierre del",
          "escenario", "flujo de caja", "flujo financiero", "tir", "margen",
          "utilidad", "presupuesto", "estado de resultados", "balance",
          "financiamiento", "banco", "q ", "q.", "millones", "gtq", "roi"]

MESES = ("enero febrero marzo abril mayo junio julio agosto septiembre "
         "setiembre octubre noviembre diciembre").split()


# --------------------------------------------------------------------------- #
#  Datos de demostración (se usan si no llega ninguna reunión real)            #
# --------------------------------------------------------------------------- #
SEED_MEETINGS = [
    {"source": "otter", "meeting_id": "demo1", "date": "2026-07-28",
     "title": "Comité semanal Boulevard Sur", "attendees": ["Paola", "Ingeniería", "Ventas"],
     "summary": "Revisión de avance Fase 2 y decisiones de socios.",
     "text": (
        "Ingeniería queda de enviar el cronograma actualizado de obra el 2026-08-04.\n"
        "Ventas se compromete a preparar la proyección de absorción para la próxima semana.\n"
        "Se decide aprobar el presupuesto de construcción de Fase 2 en Q180 millones.\n"
        "Los socios acuerdan mantener el Escenario A (vender 25/mes) por mayor TIR de accionistas.\n"
        "Contabilidad presenta el cierre contable de junio 2026: utilidad neta Q4.1 M, margen 11.4%.\n"
        "PMO revisar homologaciones pendientes con la municipalidad antes del 2026-08-08."),
     "action_items": [], "url": ""},
    {"source": "meetgeek", "meeting_id": "demo2", "date": "2026-07-21",
     "title": "Junta de socios · Flujo financiero", "attendees": ["Paola", "Socio 1", "Socio 2"],
     "summary": "Escenarios y financiamiento bancario.",
     "text": (
        "Se decide contratar el financiamiento con Banco Industrial al 8%.\n"
        "Escenario B descartado por mayor costo financiero (Q28.8 M vs Q23.1 M).\n"
        "Contabilidad enviar el cierre contable de mayo 2026 a los socios.\n"
        "Paola coordinar la reunión con G&T para comparar tasa."),
     "action_items": [{"text": "Actualizar el modelo financiero con tasa BI 8%",
                       "owner": "Contabilidad", "due": "2026-07-31"}], "url": ""},
]


# --------------------------------------------------------------------------- #
#  Helpers                                                                      #
# --------------------------------------------------------------------------- #
def _fold(s):
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


def _canon_resp(name):
    if not name:
        return ""
    key = _fold(name).strip().split()[0] if name.strip() else ""
    return RESP_ALIAS.get(key, name.strip().split(",")[0].strip().title())


def _find_owner(line):
    """Intenta deducir el responsable al inicio de la línea ('Ingeniería queda de…')."""
    m = re.match(r"\s*([A-Za-zÁÉÍÓÚÑáéíóúñ][\w. ]{1,24}?)\s+"
                 r"(queda de|se compromete|enviar|preparar|revisar|coordinar|hara|hará|entregar|asignad)",
                 line)
    if m:
        return _canon_resp(m.group(1))
    # patrón "responsable: X" o "(X)"
    m = re.search(r"responsable[:\s]+([A-Za-zÁÉÍÓÚÑáéíóúñ][\w. ]{1,24})", line, re.I)
    if m:
        return _canon_resp(m.group(1))
    return ""


def _find_due(line):
    m = re.search(r"(20\d{2})[-/ ](\d{1,2})[-/ ](\d{1,2})", line)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
        except ValueError:
            pass
    for i, mes in enumerate(MESES, 1):
        mm = re.search(rf"(\d{{1,2}})\s+de\s+{mes}\s+de\s+(20\d{{2}})", _fold(line))
        if mm:
            try:
                idx = 6 if mes == "junio" else (i if i <= 12 else 1)
                return datetime.date(int(mm.group(2)), min(idx, 12), int(mm.group(1))).isoformat()
            except ValueError:
                pass
    return ""


def _find_socios(line):
    found = [s for s in SOCIOS if _fold(s) in _fold(line)]
    if "los socios" in _fold(line) or "la junta" in _fold(line):
        found = found or ["Junta"]
    return found


def _amounts(line):
    """Extrae montos/indicadores financieros para el historial."""
    tags = []
    for m in re.finditer(r"Q\s?\.?\s?\d[\d.,]*\s?(M|millones|MM)?", line):
        tags.append(m.group(0).strip())
    for m in re.finditer(r"\b\d{1,3}(?:[.,]\d+)?\s?%", line):
        tags.append(m.group(0).strip())
    for kw in ("TIR", "margen", "utilidad", "escenario", "cierre"):
        if kw.lower() in _fold(line) and kw not in " ".join(tags):
            tags.append(kw)
    return tags[:6]


# --------------------------------------------------------------------------- #
#  Extracción                                                                   #
# --------------------------------------------------------------------------- #
def extract(meetings):
    tareas, decisiones, financiero = [], [], []
    seen_t, seen_d, seen_f = set(), set(), set()

    for mt in meetings:
        src, date = mt.get("source", ""), mt.get("date", "")
        title, url = mt.get("title", ""), mt.get("url", "")

        # 1) action items nativos (MeetGeek/Otter/Teams los pueden traer ya hechos)
        for ai in mt.get("action_items", []):
            txt = (ai.get("text") or "").strip()
            if not txt:
                continue
            k = _fold(txt)[:80]
            if k in seen_t:
                continue
            seen_t.add(k)
            tareas.append({
                "tarea": txt, "responsable": _canon_resp(ai.get("owner", "")) or "Sin asignar",
                "vence": _iso(ai.get("due", "")), "estado": "pendiente",
                "fuente": src, "reunion": title, "fecha": date, "url": url})

        # 2) líneas de la transcripción/notas
        for raw in re.split(r"[\n\r]+|(?<=[.;])\s+(?=[A-ZÁÉÍÓÚÑ])", mt.get("text", "")):
            line = raw.strip(" -•\t")
            if len(line) < 8:
                continue
            f = _fold(line)

            if any(k in f for k in KW_DECISION):
                k = f[:90]
                if k not in seen_d:
                    seen_d.add(k)
                    decisiones.append({
                        "decision": line, "socios": _find_socios(line) or ["Junta"],
                        "fecha": date, "reunion": title, "fuente": src, "url": url})

            if any(k in f for k in KW_FIN):
                k = f[:90]
                if k not in seen_f:
                    seen_f.add(k)
                    tipo = "cierre" if "cierre" in f else ("escenario" if "escenario" in f else "indicador")
                    financiero.append({
                        "tipo": tipo, "nota": line, "tags": _amounts(line),
                        "fecha": date, "reunion": title, "fuente": src, "url": url})

            # tarea: sólo si parece asignación y NO es decisión pura
            if any(k in f for k in KW_TAREA) and not any(k in f for k in KW_DECISION):
                owner = _find_owner(line)
                k = f[:90]
                if k not in seen_t and (owner or "pendiente" in f or "to do" in f or "todo" in f):
                    seen_t.add(k)
                    tareas.append({
                        "tarea": line, "responsable": owner or "Sin asignar",
                        "vence": _find_due(line), "estado": "pendiente",
                        "fuente": src, "reunion": title, "fecha": date, "url": url})

    return {"tareas": tareas, "decisiones": decisiones, "financiero": financiero}


def _iso(s):
    s = (s or "")[:10]
    try:
        datetime.date.fromisoformat(s)
        return s
    except ValueError:
        return ""


def enrich(data, meetings):
    today = datetime.date.today()
    # marca tareas vencidas
    for t in data["tareas"]:
        t["vencida"] = bool(t["vence"] and _iso(t["vence"]) and
                            datetime.date.fromisoformat(t["vence"]) < today)
    # agrupa tareas por responsable
    por_resp = {}
    for t in data["tareas"]:
        por_resp.setdefault(t["responsable"], []).append(t)
    data["por_responsable"] = dict(sorted(por_resp.items(),
                                          key=lambda kv: (-len(kv[1]), kv[0])))
    # ordena historial financiero por fecha desc
    data["financiero"].sort(key=lambda x: x.get("fecha", ""), reverse=True)
    data["decisiones"].sort(key=lambda x: x.get("fecha", ""), reverse=True)

    fuentes = sorted({m.get("source", "") for m in meetings}) or ["demo"]
    data["meta"] = {
        "reuniones": len(meetings), "fuentes": fuentes,
        "n_tareas": len(data["tareas"]),
        "n_pend": sum(1 for t in data["tareas"] if t["estado"] == "pendiente"),
        "n_venc": sum(1 for t in data["tareas"] if t.get("vencida")),
        "n_sinasig": sum(1 for t in data["tareas"] if t["responsable"] == "Sin asignar"),
        "n_decisiones": len(data["decisiones"]),
        "n_financiero": len(data["financiero"]),
        "n_responsables": len(data["por_responsable"]),
        "generado": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "es_demo": all(m.get("meeting_id", "").startswith("demo") for m in meetings) if meetings else True,
        "ultima_reunion": max((m.get("date", "") for m in meetings), default=""),
    }
    return data


def build(meetings):
    if not meetings:
        meetings = SEED_MEETINGS
    data = extract(meetings)
    return enrich(data, meetings)


def render(data):
    return TEMPLATE.replace("/*__DATA__*/",
                            "const D = " + json.dumps(data, ensure_ascii=False) + ";")


# --------------------------------------------------------------------------- #
#  Plantilla HTML (autocontenida, mismo lenguaje visual que el PMO)            #
# --------------------------------------------------------------------------- #
TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<meta http-equiv="refresh" content="3600"/>
<title>Minutas & Acciones · Boulevard Sur</title>
<style>
  :root{--navy:#0f2c4d;--navy2:#16406e;--blue:#1f73c4;--sky:#e8f0fa;--gold:#c9a227;--teal:#0f9d8c;
    --bg:#eef1f6;--card:#fff;--line:#e2e8f0;--txt:#1c2733;--muted:#67748a;--soft:#8a96a8;
    --good:#1f9d57;--bad:#d64545;--warn:#d98a1f;--shadow:0 1px 3px rgba(16,35,60,.07),0 6px 18px rgba(16,35,60,.05);}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:"Segoe UI",system-ui,-apple-system,Arial,sans-serif;background:var(--bg);color:var(--txt);line-height:1.45;font-size:13.5px}
  .wrap{max-width:1300px;margin:0 auto;padding:0 22px 55px}
  .brand{background:linear-gradient(120deg,var(--navy),var(--navy2) 55%,#1d4f86);color:#fff;padding:24px 0;box-shadow:var(--shadow)}
  .brand .wrap{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-end;gap:16px;padding-bottom:0}
  .brand .kick{color:var(--gold);font-size:11px;font-weight:800;letter-spacing:2.5px;text-transform:uppercase}
  .brand h1{font-size:23px;font-weight:700;margin-top:4px}
  .brand .sub{color:#b9cce4;font-size:12.5px;margin-top:5px}
  .brand .meta{text-align:right;font-size:12px;color:#cdddf0}
  .live{background:rgba(31,157,87,.22);border:1px solid rgba(82,217,135,.5);color:#aef0c8;font-size:11px;font-weight:700;padding:4px 11px;border-radius:20px;display:inline-flex;align-items:center;gap:6px}
  .live.warn{background:rgba(217,138,31,.18);border-color:rgba(217,138,31,.5);color:#f2cf95}
  .dot{width:7px;height:7px;border-radius:50%;background:#52d987;display:inline-block;box-shadow:0 0 0 3px rgba(82,217,135,.25)}
  .live.warn .dot{background:var(--gold);box-shadow:0 0 0 3px rgba(201,162,39,.25)}
  .nav{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap}
  .nav a{font-size:11.5px;font-weight:600;color:#cdddf0;background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.18);padding:4px 11px;border-radius:20px;text-decoration:none}
  .nav a:hover{background:rgba(255,255,255,.2)}
  section{margin-top:28px}
  h2.sec{font-size:12px;text-transform:uppercase;letter-spacing:1.1px;color:var(--navy);font-weight:800;margin-bottom:14px;display:flex;align-items:center;gap:9px}
  h2.sec::before{content:"";width:4px;height:15px;background:var(--gold);border-radius:3px}
  h2.sec .hint{font-weight:600;letter-spacing:0;text-transform:none;color:var(--soft);font-size:11.5px;margin-left:auto}
  .grid{display:grid;gap:14px}.g4{grid-template-columns:repeat(4,1fr)}.g5{grid-template-columns:repeat(5,1fr)}
  @media(max-width:1040px){.g4{grid-template-columns:repeat(2,1fr)}.g5{grid-template-columns:repeat(3,1fr)}}
  @media(max-width:560px){.g4,.g5{grid-template-columns:repeat(2,1fr)}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;box-shadow:var(--shadow)}
  .kpi .kl{color:var(--muted);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.4px}
  .kpi .kv{font-size:24px;font-weight:800;color:var(--navy);margin-top:6px;line-height:1.1}
  .kpi .ks{color:var(--soft);font-size:11.5px;margin-top:5px}
  .kpi.gold .kv{color:var(--gold)}.kpi.teal .kv{color:var(--teal)}.kpi.warn .kv{color:var(--warn)}.kpi.bad .kv{color:var(--bad)}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{text-align:left;padding:10px 11px;border-bottom:1px solid var(--line);vertical-align:top}
  thead th{color:var(--muted);font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;font-weight:700;background:#f7f9fc}
  tbody tr:hover{background:#f6f9fd}
  .resp{display:inline-block;font-size:10.5px;font-weight:700;color:var(--blue);background:var(--sky);border:1px solid #cfe0f2;padding:2px 8px;border-radius:20px}
  .resp.no{color:var(--bad);background:#fbe9e9;border-color:#f2c4c4}
  .src{display:inline-block;font-size:9.5px;font-weight:800;letter-spacing:.4px;text-transform:uppercase;padding:1px 7px;border-radius:20px;color:#fff}
  .src.teams{background:#5059c9}.src.meetgeek{background:#0f9d8c}.src.otter{background:#1f73c4}.src.demo{background:#8a96a8}
  .tg{display:inline-block;font-size:9.5px;font-weight:800;letter-spacing:.4px;text-transform:uppercase;padding:1px 7px;border-radius:20px;margin-right:6px}
  .tg.venc{color:#fff;background:var(--bad)}.tg.pend{color:#8a5a12;background:#fbe6c2;border:1px solid #eccf98}
  .dd{color:var(--bad);font-weight:800}
  .who{display:flex;flex-wrap:wrap;gap:5px;margin-top:4px}
  .chip{background:var(--sky);color:var(--blue);font-size:11px;font-weight:600;padding:3px 9px;border-radius:20px;border:1px solid #cfe0f2}
  .chip.socio{background:#fff7e6;color:#8a6d1f;border-color:#f0d79a}
  .chip.fin{background:#e9f6f3;color:#0b6b60;border-color:#bfe6df}
  .respcard h3{font-size:13.5px;color:var(--navy);display:flex;align-items:center;gap:8px}
  .respcard .cnt{font-size:10px;font-weight:800;color:var(--muted);background:#eef2f7;border:1px solid var(--line);padding:1px 7px;border-radius:20px}
  .respcard ul{list-style:none;margin-top:8px}
  .respcard li{padding:7px 0;border-bottom:1px dashed var(--line);font-size:12.5px}
  .respcard li:last-child{border:none}
  .respcard .meta{color:var(--soft);font-size:10.5px;margin-top:3px}
  .fin{display:flex;gap:12px;align-items:flex-start;padding:11px 0;border-bottom:1px dashed var(--line)}
  .fin:last-child{border:none}
  .fin .badge{flex:0 0 auto;font-size:9.5px;font-weight:800;text-transform:uppercase;letter-spacing:.4px;padding:3px 9px;border-radius:8px;color:#fff}
  .fin .badge.cierre{background:var(--navy2)}.fin .badge.escenario{background:var(--teal)}.fin .badge.indicador{background:var(--gold)}
  .fin .fdate{color:var(--soft);font-size:11px;white-space:nowrap;margin-left:auto}
  .empty{color:var(--soft);font-style:italic;font-size:12.5px;padding:14px 0}
  footer{color:var(--muted);font-size:11.5px;border-top:1px solid var(--line);padding-top:18px;margin-top:34px}
  footer code{background:#eef2f7;padding:1px 6px;border-radius:4px;color:var(--navy);font-size:11px}
  footer b{color:var(--navy)}
  .banner{background:#fff7e6;border:1px solid #f0d79a;color:#8a6d1f;border-radius:10px;padding:11px 14px;font-size:12px;margin-top:14px;display:flex;gap:9px;align-items:flex-start}
  .banner b{color:#6b520f}
</style>
</head>
<body>
  <header class="brand">
    <div class="wrap">
      <div>
        <div class="kick">Periferia Urbana · Boulevard Sur</div>
        <h1>Minutas &amp; Acciones</h1>
        <div class="sub">Tareas del equipo, decisiones de socios e historial financiero — generados de las minutas.</div>
        <div class="nav">
          <a href="dashboard.html">◧ PMO</a>
          <a href="index.html">◧ Financiero</a>
          <a href="matriz.html">◧ Matriz</a>
          <a href="minutas.html">◧ Minutas</a>
        </div>
      </div>
      <div class="meta">
        <div id="liveBadge"></div>
        <div style="margin-top:9px" id="fuentes">—</div>
        <div style="margin-top:3px">Generado <b id="gen">—</b></div>
      </div>
    </div>
  </header>

  <div class="wrap">
    <div id="demoBanner"></div>

    <section>
      <div class="grid g5" id="kpis"></div>
    </section>

    <section>
      <h2 class="sec">Asignación de tareas por responsable <span class="hint" id="thint"></span></h2>
      <div class="grid g4" id="responsables"></div>
    </section>

    <section>
      <h2 class="sec">Todas las tareas <span class="hint">extraídas de las minutas</span></h2>
      <div class="card" style="padding:0;overflow-x:auto">
        <table>
          <thead><tr><th>Tarea</th><th>Responsable</th><th>Vence</th><th>Reunión</th><th>Fuente</th></tr></thead>
          <tbody id="tareasBody"></tbody>
        </table>
      </div>
    </section>

    <section>
      <h2 class="sec">Decisiones de socios <span class="hint">acuerdos y aprobaciones</span></h2>
      <div class="card" id="decisiones"></div>
    </section>

    <section>
      <h2 class="sec">Historial financiero <span class="hint">cierres contables y escenarios</span></h2>
      <div class="card" id="financiero"></div>
    </section>

    <footer>
      <p>Este tablero se genera automáticamente de las minutas de <b>Teams</b>, <b>MeetGeek</b> y <b>Otter.ai</b>
      con <code>refresh_minutas.py</code> y se publica en GitHub Pages. Pega un solo link en Teams y siempre verá la última versión.</p>
      <p style="margin-top:6px">Las tareas también pueden crearse en <b>Microsoft Planner</b> (ver <code>SETUP-MINUTAS.md</code>).</p>
    </footer>
  </div>

<script>
/*__DATA__*/
const $ = (id) => document.getElementById(id);
const esc = (s) => (s||"").replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const M = D.meta;

// badge + cabecera
$("gen").textContent = M.generado;
$("fuentes").textContent = "Fuentes: " + M.fuentes.join(", ") + " · " + M.reuniones + " reuniones";
$("liveBadge").innerHTML = M.es_demo
  ? '<span class="live warn"><span class="dot"></span>Vista previa (datos de demo)</span>'
  : '<span class="live"><span class="dot"></span>Conectado a minutas</span>';

if (M.es_demo){
  $("demoBanner").innerHTML = '<div class="banner"><b>Vista previa.</b>&nbsp;'
    + 'Aún no hay minutas conectadas, así que se muestran datos de ejemplo. '
    + 'Configura las credenciales de Teams / MeetGeek / Otter.ai (ver SETUP-MINUTAS.md) '
    + 'para que el tablero se llene con tus reuniones reales.</div>';
}

// KPIs
const kpis = [
  {l:"Tareas", v:M.n_tareas, ks:M.n_pend+" pendientes", c:""},
  {l:"Vencidas", v:M.n_venc, ks:"requieren acción", c:M.n_venc>0?"bad":"teal"},
  {l:"Sin asignar", v:M.n_sinasig, ks:"falta responsable", c:M.n_sinasig>0?"warn":"teal"},
  {l:"Decisiones socios", v:M.n_decisiones, ks:"acuerdos registrados", c:"gold"},
  {l:"Mov. financieros", v:M.n_financiero, ks:"cierres / escenarios", c:"teal"},
];
$("kpis").innerHTML = kpis.map(k =>
  `<div class="card kpi ${k.c}"><div class="kl">${k.l}</div><div class="kv">${k.v}</div><div class="ks">${k.ks}</div></div>`
).join("");

// tarjetas por responsable
const pr = D.por_responsable || {};
$("thint").textContent = M.n_responsables + " responsables";
$("responsables").innerHTML = Object.keys(pr).length ? Object.entries(pr).map(([resp, ts]) => `
  <div class="card respcard">
    <h3><span class="resp ${resp==='Sin asignar'?'no':''}">${esc(resp)}</span><span class="cnt">${ts.length}</span></h3>
    <ul>${ts.slice(0,6).map(t => `
      <li>${t.vencida?'<span class="tg venc">Vencida</span>':'<span class="tg pend">Pend.</span>'}${esc(t.tarea)}
        <div class="meta">${t.vence?('Vence '+(t.vencida?'<span class="dd">':'')+t.vence+(t.vencida?'</span>':'')+' · '):''}${esc(t.reunion)}</div>
      </li>`).join("")}${ts.length>6?`<li class="meta">+ ${ts.length-6} más…</li>`:''}</ul>
  </div>`).join("") : '<div class="empty">No se detectaron tareas en las minutas.</div>';

// tabla de todas las tareas
$("tareasBody").innerHTML = D.tareas.length ? D.tareas.map(t => `
  <tr>
    <td>${t.vencida?'<span class="tg venc">Vencida</span>':''}${esc(t.tarea)}</td>
    <td><span class="resp ${t.responsable==='Sin asignar'?'no':''}">${esc(t.responsable)}</span></td>
    <td>${t.vence?(t.vencida?'<span class="dd">'+t.vence+'</span>':t.vence):'<span style="color:var(--soft)">—</span>'}</td>
    <td>${esc(t.reunion)}<div style="color:var(--soft);font-size:11px">${t.fecha||''}</div></td>
    <td><span class="src ${t.fuente||'demo'}">${esc(t.fuente||'demo')}</span></td>
  </tr>`).join("") : '<tr><td colspan="5" class="empty">Sin tareas.</td></tr>';

// decisiones de socios
$("decisiones").innerHTML = D.decisiones.length ? D.decisiones.map(d => `
  <div class="fin">
    <span class="badge escenario">Decisión</span>
    <div style="flex:1">
      <div>${esc(d.decision)}</div>
      <div class="who">${(d.socios||[]).map(s=>`<span class="chip socio">${esc(s)}</span>`).join("")}
        <span class="chip">${esc(d.reunion)}</span><span class="src ${d.fuente||'demo'}">${esc(d.fuente||'demo')}</span></div>
    </div>
    <span class="fdate">${d.fecha||''}</span>
  </div>`).join("") : '<div class="empty">No se detectaron decisiones de socios en las minutas.</div>';

// historial financiero
$("financiero").innerHTML = D.financiero.length ? D.financiero.map(f => `
  <div class="fin">
    <span class="badge ${f.tipo}">${esc(f.tipo)}</span>
    <div style="flex:1">
      <div>${esc(f.nota)}</div>
      <div class="who">${(f.tags||[]).map(t=>`<span class="chip fin">${esc(t)}</span>`).join("")}
        <span class="chip">${esc(f.reunion)}</span><span class="src ${f.fuente||'demo'}">${esc(f.fuente||'demo')}</span></div>
    </div>
    <span class="fdate">${f.fecha||''}</span>
  </div>`).join("") : '<div class="empty">No se detectaron cierres ni escenarios financieros en las minutas.</div>';
</script>
</body>
</html>"""


def main():
    meetings = SEED_MEETINGS
    if "--stdin" in sys.argv:
        try:
            meetings = json.load(sys.stdin) or SEED_MEETINGS
        except Exception:
            meetings = SEED_MEETINGS
    data = build(meetings)
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "minutas.html")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(render(data))
    m = data["meta"]
    print(f"[ok] minutas.html generado · {m['reuniones']} reuniones · "
          f"{m['n_tareas']} tareas · {m['n_decisiones']} decisiones · "
          f"{m['n_financiero']} mov. financieros" + (" · DEMO" if m["es_demo"] else ""))


if __name__ == "__main__":
    main()
