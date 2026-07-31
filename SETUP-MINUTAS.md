# Minutas & Acciones — puesta en marcha

Este módulo lee las **minutas de reuniones** de **Teams**, **MeetGeek** y
**Otter.ai**, y genera `minutas.html` con tres paneles:

1. **Asignación de tareas** — quién hace qué, para cuándo, de qué reunión salió.
2. **Decisiones de socios** — acuerdos y aprobaciones tomadas por los socios.
3. **Historial financiero** — cierres contables y escenarios financieros mencionados.

Además, opcionalmente, **crea esas tareas en Microsoft Planner**.

```
┌ Teams / MeetGeek / Otter.ai ┐   ┌ GitHub Actions (diario) ┐   ┌ GitHub Pages ┐
│  minutas / transcripciones   │→ │ refresh_minutas.py        │→ │ minutas.html │→ Link en Teams
│                              │   │  · junta reuniones        │   └──────────────┘
└──────────────────────────────┘   │  · extrae tareas/decis.   │        │
                                    │  · genera el HTML         │   ┌ Planner ┐
                                    │  · (opcional) → Planner ──┼─→ │ tareas  │
                                    └───────────────────────────┘   └─────────┘
```

Todo degrada con elegancia: **cada fuente que no tenga credenciales se salta
sola**, y si no hay ninguna, el tablero muestra datos de demo. No hace falta
configurarlo todo de una vez.

---

## Probar sin credenciales

```bash
python3 build_minutas.py     # genera minutas.html con datos de ejemplo
```

Abre `minutas.html` en el navegador para ver el diseño con datos de demo.

---

## Fuente 1 · Teams / SharePoint (recomendada, la más simple)

En vez de leer transcripciones crudas de Teams (que requieren permisos
especiales), lo más robusto es tener **una carpeta de SharePoint** donde caigan
las minutas: el resumen que exporta MeetGeek/Otter, el `.docx` de notas, o el
`.vtt` de la transcripción de Teams. El módulo lee `.txt`, `.md`, `.vtt`, `.docx`
y `.csv` de esa carpeta y toma cada archivo como una reunión.

Reutiliza la **misma app de Azure** del dashboard PMO (permiso `Sites.Read.All`).
Sólo hay que decirle **qué carpeta** leer:

1. Crea una carpeta, p. ej. `Boulevard Sur › Minutas`.
2. Obtén su `GRAPH_MINUTAS_FOLDER_ID` (el id del item de esa carpeta en Graph).
   La forma rápida: abre la carpeta en el navegador y usa Graph Explorer, o pide
   a IT el id. Si la carpeta está en **otra biblioteca**, define también
   `GRAPH_MINUTAS_DRIVE_ID`.
3. Guarda estos Secrets del repositorio (además de los `AZURE_*` que ya existen):

| Secret | Valor |
|---|---|
| `GRAPH_MINUTAS_FOLDER_ID` | id de la carpeta de minutas |
| `GRAPH_MINUTAS_DRIVE_ID` | *(opcional)* id de la biblioteca si no es la del PMO |

> **Transcripción automática de Teams:** si quieres leer las transcripciones que
> Teams genera solo (sin carpeta intermedia), la app de Azure necesita el permiso
> `OnlineMeetingTranscript.Read.All` con una *access policy* de aplicación. Es
> más laborioso; la carpeta de SharePoint cubre el 90 % de los casos.

---

## Fuente 2 · MeetGeek (API)

1. En MeetGeek: **Settings → API** → genera un **API key**.
2. Guarda el Secret:

| Secret | Valor |
|---|---|
| `MEETGEEK_API_KEY` | tu API key de MeetGeek |

El módulo trae las reuniones recientes, su **resumen** y sus **highlights**
(los marcados como *action item* se convierten en tareas).

---

## Fuente 3 · Otter.ai

Otter no publica una API oficial estable; el módulo usa el mismo endpoint que la
web de Otter, autenticándose con tu usuario:

| Secret | Valor |
|---|---|
| `OTTER_EMAIL` | correo de la cuenta de Otter |
| `OTTER_PASSWORD` | contraseña de la cuenta de Otter |

> Recomendación: usa una cuenta de servicio dedicada. En una sesión interactiva
> de Claude también puedes leer Otter por el **conector MCP de Otter.ai** sin
> guardar la contraseña.

---

## Crear las tareas en Microsoft Planner (opcional)

Por defecto el módulo **sólo genera el tablero** (no escribe nada). Para que
además cree las tareas en Planner:

### 1) Dar permiso de escritura a la app de Azure
En la app de Azure → **API permissions** → **Microsoft Graph** →
**Application permissions** → agrega **`Group.ReadWrite.All`** (o
`Tasks.ReadWrite`) → **Grant admin consent**.

### 2) Identificar el plan destino
Abre el **Planner** del equipo y copia el `planId` de la URL (o pídelo a IT).
Opcionalmente un `bucketId`; si no lo das, se usa el primer bucket del plan.

### 3) Guardar los Secrets

| Secret | Valor |
|---|---|
| `PLANNER_ENABLED` | `1` para activar la escritura |
| `PLANNER_PLAN_ID` | id del plan de Planner |
| `PLANNER_BUCKET_ID` | *(opcional)* bucket destino |
| `PLANNER_ASSIGNEES` | *(opcional)* JSON `{"Contabilidad":"guid-usuario","Paola":"guid-usuario"}` para asignar |

La escritura es **idempotente**: no duplica una tarea cuyo título ya exista en el
plan, así que se puede correr todos los días sin problema. Las tareas sin
responsable en el mapa se crean **sin asignar**.

---

## Ajustar el equipo y las heurísticas

En `build_minutas.py`, arriba del todo:

- `SOCIOS` — nombres de los socios (sus líneas alimentan *Decisiones de socios*).
- `RESP_ALIAS` — alias → nombre canónico de responsables (`"conta" → "Contabilidad"`).
- `KW_TAREA` / `KW_DECISION` / `KW_FIN` — palabras clave que clasifican cada línea
  como tarea, decisión o movimiento financiero.

---

## Cada cuánto se actualiza

El módulo corre dentro del workflow **`refresh-dashboard.yml`** (diario 06:00 GT y
cuando lo dispares a mano en la pestaña **Actions**). Publica `minutas.html` en
GitHub Pages junto con el resto:

```
https://paolamp91.github.io/finanzas/minutas.html
```

Pega ese link en Teams (como pestaña *Website* o como mensaje) y siempre mostrará
la última versión.
