#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
refresh_minutas.py
==================
Orquestador del módulo Minutas & Acciones:

  1. Junta reuniones de las fuentes CONFIGURADAS (Teams, MeetGeek, Otter.ai).
  2. Extrae tareas, decisiones de socios e historial financiero.
  3. Genera `minutas.html`.
  4. (Opcional) crea las tareas en Microsoft Planner si PLANNER_ENABLED=1.

Si no hay ninguna fuente configurada, genera `minutas.html` con datos de demo
para que el sitio no quede vacío. Cada fuente sin credenciales se salta sin
interrumpir el proceso.

Uso:
    python3 refresh_minutas.py
"""
import os, sys, json

import minutas_sources as srcs
import build_minutas
import push_to_planner


def main():
    here = os.path.dirname(os.path.abspath(__file__))

    meetings = srcs.fetch_all()
    for s in srcs.skipped:
        print("[info] fuente saltada — " + s)

    if meetings:
        print(f"[ok] {len(meetings)} reuniones reales recopiladas de "
              f"{sorted({m['source'] for m in meetings})}")
    else:
        print("[info] ninguna fuente configuró credenciales; se usan datos de demo.")

    data = build_minutas.build(meetings)   # cae a SEED si meetings está vacío
    html = build_minutas.render(data)
    out = os.path.join(here, "minutas.html")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(html)
    m = data["meta"]
    print(f"[ok] minutas.html · {m['n_tareas']} tareas · {m['n_decisiones']} decisiones · "
          f"{m['n_financiero']} mov. financieros" + (" · DEMO" if m["es_demo"] else ""))

    # Escritura opcional a Planner (sólo tareas reales, nunca las de demo)
    if not m["es_demo"]:
        try:
            push_to_planner.push(data["tareas"])
        except SystemExit as e:
            print(f"[planner] no se escribió (auth): {e}")
        except Exception as e:
            print(f"[planner] error inesperado: {type(e).__name__}: {e}")
    elif push_to_planner.enabled():
        print("[planner] activado, pero los datos son de demo: no se crea nada en Planner.")


if __name__ == "__main__":
    main()
