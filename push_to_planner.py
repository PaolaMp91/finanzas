#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
push_to_planner.py
==================
Crea/actualiza tareas en Microsoft Planner a partir de las tareas extraídas de
las minutas. **Desactivado por defecto**: sólo escribe si PLANNER_ENABLED=1 y
existe PLANNER_PLAN_ID. Es idempotente: no duplica una tarea cuyo título ya
exista en el plan.

Requiere que la app de Azure tenga permiso de aplicación de ESCRITURA
(Group.ReadWrite.All o Tasks.ReadWrite) con consentimiento de administrador,
ADEMÁS del de lectura. El token es el mismo; los permisos se otorgan en Azure.

Variables de entorno:
  PLANNER_ENABLED     "1" para activar la escritura (por defecto NO escribe)
  PLANNER_PLAN_ID     id del plan de Planner donde crear las tareas
  PLANNER_BUCKET_ID   (opcional) bucket destino; si falta, usa el primero del plan

Mapeo de responsable → usuario de Planner (para asignar): PLANNER_ASSIGNEES,
un JSON como {"Contabilidad":"aad-user-guid","Paola":"aad-user-guid"}.
Si un responsable no está en el mapa, la tarea se crea SIN asignar.
"""
import os, json, urllib.error

import graph_common as gc


def enabled():
    return gc.clean("PLANNER_ENABLED") == "1" and bool(gc.clean("PLANNER_PLAN_ID"))


def _assignees_map():
    try:
        return json.loads(gc.clean("PLANNER_ASSIGNEES") or "{}")
    except Exception:
        return {}


def _existing_titles(token, plan_id):
    h = {"Authorization": f"Bearer {token}"}
    titles = set()
    url = f"{gc.GRAPH}/planner/plans/{plan_id}/tasks?$top=200"
    while url:
        data = gc.http_json(url, headers=h)
        for t in data.get("value", []):
            titles.add((t.get("title") or "").strip().lower())
        url = data.get("@odata.nextLink")
    return titles


def _first_bucket(token, plan_id):
    h = {"Authorization": f"Bearer {token}"}
    data = gc.http_json(f"{gc.GRAPH}/planner/plans/{plan_id}/buckets", headers=h)
    vals = data.get("value", [])
    return vals[0]["id"] if vals else None


def push(tareas):
    """Crea en Planner las tareas que aún no existan. Devuelve (creadas, saltadas)."""
    if not enabled():
        print("[planner] desactivado (PLANNER_ENABLED != 1 o falta PLANNER_PLAN_ID). "
              "No se escribió nada.")
        return 0, len(tareas)

    plan_id = gc.clean("PLANNER_PLAN_ID")
    bucket = gc.clean("PLANNER_BUCKET_ID")
    amap = _assignees_map()
    token = gc.get_token()
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    if not bucket:
        bucket = _first_bucket(token, plan_id)
        if not bucket:
            print("[planner] el plan no tiene buckets; crea uno o define PLANNER_BUCKET_ID.")
            return 0, len(tareas)

    existing = _existing_titles(token, plan_id)
    creadas, saltadas = 0, 0
    for t in tareas:
        title = (t.get("tarea") or "").strip()[:255]
        if not title or title.lower() in existing:
            saltadas += 1
            continue
        body = {"planId": plan_id, "bucketId": bucket, "title": title}
        due = t.get("vence")
        if due:
            body["dueDateTime"] = f"{due}T17:00:00Z"
        uid = amap.get(t.get("responsable", ""))
        if uid:
            body["assignments"] = {uid: {"@odata.type": "#microsoft.graph.plannerAssignment",
                                         "orderHint": " !"}}
        try:
            gc.http_json(f"{gc.GRAPH}/planner/tasks", method="POST",
                         data=json.dumps(body).encode(), headers=h)
            existing.add(title.lower())
            creadas += 1
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "replace")[:200]
            print(f"[planner] no se pudo crear «{title[:40]}…» (HTTP {e.code}): {detail}")
            saltadas += 1
    print(f"[planner] {creadas} tareas creadas, {saltadas} ya existían o se omitieron.")
    return creadas, saltadas
