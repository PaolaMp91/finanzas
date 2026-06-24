#!/usr/bin/env python3
"""
build_from_matrix.py
--------------------
Se conecta a la MATRIZ (archivo maestro Excel del flujo financiero) y genera
el dashboard `matriz.html` con sus valores reales.

Uso:
    python3 build_from_matrix.py "<ruta_al_excel>.xlsx"

Para refrescar con nuevos valores: descargue la última versión de la matriz desde
Teams (Periferia Urbana › Boulevard Sur › Flujo Financiero) y vuelva a ejecutar.
Hojas leídas: "Dashboard (2)" (Ejecutado vs Proyectado) y "RESUMEN GENERAL".
"""
import sys, json, datetime
import openpyxl

DEFAULT_XLSX = "/root/.claude/uploads/c5b42dfd-6333-5973-96f2-aa28d9aa4a67/76be16e9-PU_PROJECT_94M_Escenario_BI_31052026_12_nvls_1.xlsx"

def num(v):
    return float(v) if isinstance(v, (int, float)) else None

def read_sheet(ws, max_row, max_col):
    """Lee una hoja en una sola pasada -> dict {fila(1-index): tupla}."""
    out = {}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True), 1):
        out[i] = row
    return out

def extract(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    R = read_sheet(wb["Dashboard (2)"], 36, 9)

    # Dashboard(2) índices (0-based): 2=rubro 3=proyGTQ 4=USD 5=ejecutado 6=%ejec 7=por_ejec
    control = []
    for i in range(15, 32):
        row = R.get(i)
        if not row: continue
        rubro = str(row[2]).strip() if row[2] else None
        proy = num(row[3]); ejec = num(row[5])
        if rubro and proy is not None and abs(proy) > 0:
            p = abs(proy); e = abs(ejec) if ejec else 0.0
            control.append({"rubro": rubro, "proyectado": p, "ejecutado": e,
                            "pct": (e / p if p else 0)})
    ventas = num(R[33][3]) or 0
    utilidad = num(R[34][3]) or 0
    totals = {
        "costo_total": abs(num(R[32][3]) or 0),
        "ejecutado_total": abs(num(R[32][5]) or 0),
        "ventas": ventas,
        "utilidad": utilidad,
        "margen": (utilidad / ventas if ventas else 0),
    }
    corte = "Enero 2026"
    actualizado = ""
    for i in range(1, 7):
        for v in (R.get(i) or ()):
            if isinstance(v, datetime.datetime):
                actualizado = v.strftime("%d/%m/%Y"); break
        if actualizado: break

    # ---- RESUMEN GENERAL: métricas consolidadas (base "fechas actuales") ----
    rg = read_sheet(wb["RESUMEN GENERAL"], 32, 12)
    def col(i, c): return num(rg[i][c]) or 0
    # col4=F1, col6=F2, col8=GENERAL BI, col10=GENERAL G&T
    metrics = {
        "F1":  {"ventas": col(27,4), "util": col(28,4), "margen": col(29,4), "tir_proy": col(30,4), "tir_acc": col(31,4)},
        "F2":  {"ventas": col(27,6), "util": col(28,6), "margen": col(29,6), "tir_proy": col(30,6), "tir_acc": col(31,6)},
        "BI":  {"ventas": col(27,8), "util": col(28,8), "margen": col(29,8), "tir_proy": col(30,8), "tir_acc": col(31,8)},
        "GT":  {"ventas": col(27,10),"util": col(28,10),"margen": col(29,10),"tir_proy": col(30,10),"tir_acc": col(31,10)},
    }
    return {"corte": corte, "actualizado": actualizado, "control": control,
            "totals": totals, "metrics": metrics, "fuente": path.split("/")[-1]}


def render(data):
    J = json.dumps(data, ensure_ascii=False)
    return TEMPLATE.replace("/*DATA*/", "const M = " + J + ";")


TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Matriz Financiera · Boulevard Sur Club Residencial</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{--navy:#0f2c4d;--navy2:#16406e;--blue:#1f73c4;--sky:#e8f0fa;--gold:#c9a227;--teal:#0f9d8c;
    --bg:#eef1f6;--card:#fff;--line:#e2e8f0;--txt:#1c2733;--muted:#67748a;--soft:#8a96a8;
    --good:#1f9d57;--bad:#d64545;--warn:#d98a1f;--shadow:0 1px 3px rgba(16,35,60,.07),0 6px 18px rgba(16,35,60,.05);}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:"Segoe UI",system-ui,Arial,sans-serif;background:var(--bg);color:var(--txt);line-height:1.45;font-size:13.5px}
  .wrap{max-width:1280px;margin:0 auto;padding:0 22px 50px}
  .nav{background:var(--navy);padding:9px 0}.nav .wrap{display:flex;gap:8px;padding-bottom:0}
  .nav a{color:#b9cce4;text-decoration:none;font-size:12.5px;font-weight:600;padding:6px 13px;border-radius:7px}
  .nav a.on{background:rgba(255,255,255,.13);color:#fff}.nav a:hover{color:#fff}
  .brand{background:linear-gradient(120deg,var(--navy),var(--navy2) 55%,#1d4f86);color:#fff;padding:24px 0;box-shadow:var(--shadow)}
  .brand .wrap{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-end;gap:14px;padding-bottom:0}
  .brand .kick{color:var(--gold);font-size:11px;font-weight:800;letter-spacing:2.5px;text-transform:uppercase}
  .brand h1{font-size:24px;font-weight:700;margin-top:3px}
  .brand .sub{color:#b9cce4;font-size:12.5px;margin-top:4px}
  .brand .live{background:rgba(31,157,87,.25);border:1px solid rgba(82,217,135,.5);color:#aef0c8;font-size:11px;font-weight:700;padding:4px 11px;border-radius:20px;display:inline-flex;align-items:center;gap:6px}
  .dot{width:7px;height:7px;border-radius:50%;background:#52d987;display:inline-block;box-shadow:0 0 0 3px rgba(82,217,135,.25)}
  section{margin-top:26px}
  h2.sec{font-size:12px;text-transform:uppercase;letter-spacing:1.1px;color:var(--navy);font-weight:800;margin-bottom:13px;display:flex;align-items:center;gap:9px}
  h2.sec::before{content:"";width:4px;height:15px;background:var(--gold);border-radius:3px}
  .grid{display:grid;gap:14px}.g5{grid-template-columns:repeat(5,1fr)}.g4{grid-template-columns:repeat(4,1fr)}.g3{grid-template-columns:repeat(3,1fr)}.g2{grid-template-columns:1.25fr 1fr}
  @media(max-width:1000px){.g5,.g4{grid-template-columns:repeat(2,1fr)}.g3{grid-template-columns:1fr}.g2{grid-template-columns:1fr}}
  @media(max-width:560px){.g5,.g4{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;box-shadow:var(--shadow)}
  .kpi .kl{color:var(--muted);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.4px}
  .kpi .kv{font-size:23px;font-weight:800;color:var(--navy);margin-top:6px}
  .kpi .ks{color:var(--soft);font-size:11px;margin-top:4px}
  .kpi.hero{background:linear-gradient(150deg,var(--navy),var(--navy2));border:none}
  .kpi.hero .kl{color:#a9c2e0}.kpi.hero .kv{color:#fff}.kpi.hero .ks{color:#94aece}
  .kpi.gold .kv{color:var(--gold)}.kpi.teal .kv{color:var(--teal)}
  table{width:100%;border-collapse:collapse;font-size:12.5px}
  th,td{text-align:right;padding:7px 9px;border-bottom:1px solid var(--line)}
  th:first-child,td:first-child{text-align:left}
  thead th{color:var(--muted);font-size:10px;text-transform:uppercase;letter-spacing:.5px;font-weight:700;background:#f7f9fc;position:sticky;top:0}
  tbody tr:hover{background:#f6f9fd}
  tfoot td{font-weight:800;border-top:2px solid var(--navy);border-bottom:none;color:var(--navy)}
  .prog{position:relative;height:16px;background:#eef2f7;border-radius:8px;overflow:hidden;min-width:90px}
  .prog>div{position:absolute;left:0;top:0;bottom:0;background:linear-gradient(90deg,var(--teal),#3bbfa9);border-radius:8px}
  .prog span{position:absolute;right:6px;top:0;font-size:10px;font-weight:700;color:var(--navy);line-height:16px}
  .chart-box{position:relative;height:260px}
  .pos{color:var(--good);font-weight:700}.neg{color:var(--bad);font-weight:700}
  footer{color:var(--muted);font-size:11px;border-top:1px solid var(--line);padding-top:16px;margin-top:26px}
  footer code{background:#eef2f7;padding:1px 6px;border-radius:4px;color:var(--navy)}
</style></head>
<body>
  <div class="nav"><div class="wrap">
    <a href="index.html">Dashboard operativo</a>
    <a href="brief.html">Investment Brief</a>
    <a href="matriz.html" class="on">Matriz · Ejecutado vs Proyectado</a>
  </div></div>

  <div class="brand"><div class="wrap">
    <div>
      <div class="kick">Conectado a la Matriz Financiera</div>
      <h1>Boulevard Sur Club Residencial</h1>
      <div class="sub" id="sub"></div>
    </div>
    <div style="text-align:right">
      <span class="live"><span class="dot"></span> <span id="liveTxt"></span></span>
    </div>
  </div></div>

  <div class="wrap">
    <section>
      <h2 class="sec">Indicadores Consolidados de la Matriz <span style="font-weight:600;color:var(--muted);text-transform:none;letter-spacing:0">· Banco Industrial 8%</span></h2>
      <div class="grid g5" id="kpis"></div>
    </section>

    <section>
      <h2 class="sec">Cuadro de Control Financiero · Ejecutado vs. Proyectado</h2>
      <div class="grid g2">
        <div class="card" style="overflow:auto;max-height:560px">
          <table>
            <thead><tr><th>Rubro</th><th>Proyectado (Q)</th><th>Ejecutado (Q)</th><th>Avance</th></tr></thead>
            <tbody id="ctrlBody"></tbody>
            <tfoot id="ctrlFoot"></tfoot>
          </table>
        </div>
        <div class="card">
          <div class="chart-box" style="height:300px"><canvas id="chTop"></canvas></div>
          <div id="execSummary" style="margin-top:14px"></div>
        </div>
      </div>
    </section>

    <section>
      <h2 class="sec">Rentabilidad por Fase y por Banco · valores de la matriz</h2>
      <div class="grid g2">
        <div class="card">
          <table>
            <thead><tr><th>Métrica</th><th>Fase 1</th><th>Fase 2</th><th>General BI</th><th>General G&amp;T</th></tr></thead>
            <tbody id="metBody"></tbody>
          </table>
        </div>
        <div class="card"><div class="chart-box"><canvas id="chTir"></canvas></div></div>
      </div>
    </section>

    <footer id="foot"></footer>
  </div>

<script>
/*DATA*/
const fmtQ=v=>"Q "+Math.round(v).toLocaleString("es-GT");
const fmtM=v=>"Q "+(v/1e6).toFixed(1)+" M";
const pct=v=>(v*100).toFixed(1)+"%";

document.getElementById("sub").textContent="Cuadro de control al "+M.corte+" · Matriz: "+M.fuente;
document.getElementById("liveTxt").textContent="Actualizado "+M.actualizado;

// KPIs
const bi=M.metrics.BI, f2=M.metrics.F2;
const kpis=[
  ["hero","Ventas Totales",fmtM(M.totals.ventas),"F1+F2 · matriz"],
  ["hero","Costo Total",fmtM(M.totals.costo_total),"Inversión proyectada"],
  ["gold","TIR Accionistas",pct(bi.tir_acc),"G&T "+pct(M.metrics.GT.tir_acc)],
  ["teal","TIR Proyecto",pct(bi.tir_proy),"G&T "+pct(M.metrics.GT.tir_proy)],
  ["","Utilidad / Margen",fmtM(M.totals.utilidad),"Margen "+pct(M.totals.margen)],
];
document.getElementById("kpis").innerHTML=kpis.map(k=>
  `<div class="card kpi ${k[0]}"><div class="kl">${k[1]}</div><div class="kv">${k[2]}</div><div class="ks">${k[3]}</div></div>`).join("");

// Control table
const ctrl=[...M.control].sort((a,b)=>b.proyectado-a.proyectado);
document.getElementById("ctrlBody").innerHTML=ctrl.map(r=>{
  const p=Math.min(r.pct,1)*100;
  return `<tr><td>${r.rubro}</td><td>${fmtQ(r.proyectado)}</td><td>${fmtQ(r.ejecutado)}</td>
    <td><div class="prog"><div style="width:${p.toFixed(0)}%"></div><span>${(r.pct*100).toFixed(0)}%</span></div></td></tr>`;
}).join("");
const tEjec=M.totals.ejecutado_total, tProy=M.totals.costo_total;
document.getElementById("ctrlFoot").innerHTML=
  `<tr><td>COSTO TOTAL</td><td>${fmtQ(tProy)}</td><td>${fmtQ(tEjec)}</td><td>${(tEjec/tProy*100).toFixed(1)}%</td></tr>`;
document.getElementById("execSummary").innerHTML=
  `<div style="font-size:12.5px;color:var(--muted)">Avance global de ejecución del costo del proyecto:
   <b style="color:var(--navy);font-size:15px">${(tEjec/tProy*100).toFixed(1)}%</b> · Ejecutado ${fmtM(tEjec)} de ${fmtM(tProy)}.
   Por ejecutar: <b>${fmtM(tProy-tEjec)}</b>.</div>`;

// Metrics table
const order=[["Ventas","ventas","M"],["Utilidad neta","util","M"],["Margen neto","margen","P"],["TIR proyecto","tir_proy","P"],["TIR accionistas","tir_acc","P"]];
const cols=["F1","F2","BI","GT"];
document.getElementById("metBody").innerHTML=order.map(o=>{
  const cells=cols.map(c=>{const v=M.metrics[c][o[1]];
    const t=o[2]==="M"?fmtM(v):pct(v);
    const cls=(o[1]==="util"||o[1]==="margen")&&v<0?"neg":(v>0&&o[2]==="P"?"pos":"");
    return `<td class="${cls}">${t}</td>`;}).join("");
  return `<tr><td>${o[0]}</td>${cells}</tr>`;
}).join("");

// Charts
Chart.defaults.color="#67748a";Chart.defaults.borderColor="#e2e8f0";Chart.defaults.font.family="Segoe UI, system-ui, sans-serif";
const top=[...M.control].sort((a,b)=>b.proyectado-a.proyectado).slice(0,7);
new Chart(document.getElementById("chTop"),{type:"bar",
  data:{labels:top.map(r=>r.rubro.length>18?r.rubro.slice(0,17)+"…":r.rubro),
    datasets:[{label:"Proyectado",data:top.map(r=>r.proyectado/1e6),backgroundColor:"#cdd7e2",borderRadius:4},
              {label:"Ejecutado",data:top.map(r=>r.ejecutado/1e6),backgroundColor:"#0f9d8c",borderRadius:4}]},
  options:{indexAxis:"y",plugins:{legend:{position:"bottom"},title:{display:true,text:"Top rubros: proyectado vs ejecutado (Q M)",color:"#0f2c4d",font:{size:13,weight:"600"}}},
    scales:{x:{ticks:{callback:v=>"Q"+v+"M"}}}}});
new Chart(document.getElementById("chTir"),{type:"bar",
  data:{labels:["Fase 1","Fase 2","General BI","General G&T"],
    datasets:[{label:"TIR Accionistas",data:cols.map(c=>M.metrics[c].tir_acc*100),backgroundColor:"#c9a227",borderRadius:5},
              {label:"TIR Proyecto",data:cols.map(c=>M.metrics[c].tir_proy*100),backgroundColor:"#16406e",borderRadius:5}]},
  options:{plugins:{legend:{position:"bottom"},title:{display:true,text:"TIR por fase y banco (%)",color:"#0f2c4d",font:{size:13,weight:"600"}}},
    scales:{y:{beginAtZero:true,ticks:{callback:v=>v+"%"}}}}});

document.getElementById("foot").innerHTML=
  "<b>Generado desde la matriz</b> <code>"+M.fuente+"</code> mediante <code>build_from_matrix.py</code> (hojas \"Dashboard (2)\" y \"RESUMEN GENERAL\"). "+
  "Para refrescar con nuevos valores, descargue la última versión de la matriz desde Teams › Periferia Urbana › Boulevard Sur › Flujo Financiero y vuelva a ejecutar el script. Cifras en quetzales · Banco Industrial 8%.";
</script>
</body></html>"""


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    data = extract(path)
    html = render(data)
    with open("matriz.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("OK matriz.html generado desde:", data["fuente"])
    print("  rubros:", len(data["control"]),
          "| ventas:", round(data["totals"]["ventas"]),
          "| utilidad:", round(data["totals"]["utilidad"]),
          "| TIR acc BI:", round(data["metrics"]["BI"]["tir_acc"]*100,2), "%")
