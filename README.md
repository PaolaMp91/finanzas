# Dashboard Financiero · Boulevard Sur Club Residencial

Dashboard interactivo (un solo archivo HTML) construido sobre los **datos reales del archivo maestro**
`PU PROJECT 94M Escenario BI 31-05-2026 12 nvls.xlsx` (Periferia Urbana · Mixco, Guatemala).

## Cómo usarlo
Abra **`index.html`** en cualquier navegador (doble clic). Internet solo se requiere la primera vez para Chart.js (CDN).

## Datos reales del proyecto (Excel · Banco Industrial 8%)
- Ventas totales: **Q449.7 M** (F1: Q144.1M · F2: Q305.7M)
- Inversión total: **Q398.6 M** · Utilidad neta: **Q51.1 M** · Margen: **11.37%**
- TIR proyecto: **12.53%** · TIR accionistas: **29.18%** (G&T: 13.12% / 30.91%)
- 648 apartamentos (F1: 216 · F2: 432) · 898 parqueos (F1: 321 · F2: 577)

La Fase 1 opera con margen negativo (−4.79%, alta carga de infraestructura); la Fase 2 concentra la rentabilidad
(margen 18.98%, TIR accionistas 63.6%).

## Comparativa estratégica (Fase 2)
Modelo pro-forma mensual **calibrado contra el Excel** (la base reproduce TIR proyecto 21.7% vs 21.9% y TIR
accionistas 65.7% vs 63.6% del archivo). Aplica los cambios solicitados:

- Presupuesto de construcción → **Q180 M**
- Promoción → **1% del presupuesto total**
- **Sin ingreso de bodegas** en Fase 2
- **Entregas empatadas con la finalización de la obra**

| Métrica | A · Vender 25/mes (rápido) | B · Vender 14/mes (+3% anual) |
|---|---|---|
| TIR Accionistas | **52.4%** | 41.7% |
| TIR Proyecto | **15.7%** | 14.2% |
| Utilidad neta | Q29.3 M | **Q30.5 M** |
| Margen | 9.6% | 9.7% |
| Ventas | Q305.7 M | Q313.4 M |
| Costo financiero | **Q23.1 M** | Q28.8 M |
| Duración de obra | 18 meses | 26 meses |
| 1ª entrega | Oct 2028 | Jun 2029 |

**Recomendación: Escenario A.** Mayor TIR de accionistas (+10.7 pp), entrega ≈8 meses antes y menor costo
financiero (−Q5.7 M). El alza de precio del 3% anual de B solo aporta Q1.2 M más de utilidad absoluta,
insuficiente para compensar el mayor tiempo, interés y exposición al riesgo.

## Metodología del modelo
`model.py` (incluido) construye un flujo de caja mensual de Fase 2: ventas por absorción, enganche 10% en cuotas,
hipoteca 90% a escrituración (3 meses tras entrega), construcción en S-curve, comisiones, promoción, financiamiento
bancario al 8% e ISR. La TIR se calcula mensual y se anualiza. La base se calibra contra el Estado de Resultados
del Excel antes de proyectar los escenarios.

## Vista conectada a la matriz (`matriz.html`)
Dashboard generado **directamente desde la matriz** (Excel maestro) con el script `build_from_matrix.py`:

```bash
python3 build_from_matrix.py "ruta/al/PU PROJECT 94M Escenario BI ... .xlsx"
```

Lee las hojas `Dashboard (2)` (Cuadro de Control **Ejecutado vs Proyectado**) y `RESUMEN GENERAL`
(rentabilidad por fase y banco) y produce `matriz.html` con los valores reales: ventas Q449.7M,
utilidad Q51.1M, margen 11.37%, TIR accionistas 29.18% (BI) / 30.91% (G&T), y el avance de ejecución
del costo por rubro. **Para refrescar con nuevos valores**, descargue la última versión de la matriz
desde Teams y vuelva a ejecutar el script.

## Secciones del dashboard
Resumen consolidado real · Comparativa estratégica A vs B con recomendación · Gráficas de TIR y utilidad ·
Unidades por fase · Presupuesto de inversión F2 (con los cambios) · Cronograma y comparativo bancario BI/G&T ·
Flujo de costos de construcción mensual proyectado vs. real (editable, por escenario).

## Fuente
Microsoft Teams › Periferia Urbana › Canal **Boulevard Sur** › carpeta **Flujo Financiero** →
`PU PROJECT 94M Escenario BI 31-05-2026 12 nvls.xlsx` (hojas RESUMEN GENERAL, ER F2, SUPUESTOS F2, INDICES, INCREMENTO F2).

---

# KPI de Planner (`planner-kpi.html`)

Dashboard de **cumplimiento de tareas del plan "Periferia Urbana" de Microsoft Planner**, conectado en vivo
a Microsoft Graph. No usa exportaciones ni copias: cada vez que se abre, lee las tareas directamente de
Planner con la sesión de Microsoft del usuario, y se refresca solo cada 5 minutos mientras está abierto.

**Filtros:** usuario · período · estado (realizadas en fecha estimada / realizadas fuera de fecha /
pendientes vencidas / pendientes en plazo).

**KPIs:** % de cumplimiento de entrega, entregadas a tiempo, entregadas tarde, vencidas sin entregar
y avance de la operación — en total, por usuario (gráfica y tabla) y con el detalle tarea por tarea.

> El cumplimiento se mide **contra la fecha de entrega** (fecha de vencimiento de Planner):
> **a tiempo** = entrega real ≤ fecha de entrega (por día), y
> **% cumplimiento = a tiempo ÷ exigibles**, donde exigibles = entregadas con fecha + vencidas sin
> entregar. Las entregadas sin fecha de entrega se listan aparte y no afectan el %.

## Puesta en marcha (una sola vez)

### 1. Registrar la aplicación en Azure (lo hace TI o un administrador, ~5 min)
1. [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **Registros de aplicaciones** → **Nuevo registro**.
2. Nombre: `Dashboard KPI Planner` · Cuentas: **solo este directorio organizativo**.
3. URI de redirección: tipo **Aplicación de página única (SPA)** → la URL pública de la página
   (p. ej. `https://paolamp91.github.io/finanzas/planner-kpi.html`).
4. **Permisos de API** → Microsoft Graph → **Delegados**: `User.Read`, `Tasks.Read`, `User.ReadBasic.All`
   → **Conceder consentimiento de administrador**.
5. Copiar el **Id. de aplicación (cliente)** y el **Id. de directorio (inquilino)**.

### 2. Publicar la página
En GitHub → repositorio `finanzas` → **Settings → Pages** → Source: rama `main` (carpeta `/`).
La página queda en `https://paolamp91.github.io/finanzas/planner-kpi.html`.

### 3. Dejarlo en el canal **General Administrativo** de Teams
1. Abrir el equipo → canal **General Administrativo** → **+** (Agregar pestaña) → **Sitio web**.
2. Nombre: `KPI Planner` · URL (con la configuración incluida, para que nadie tenga que configurar nada):

   ```
   https://paolamp91.github.io/finanzas/planner-kpi.html?clientId=<CLIENT_ID>
   ```

   El tenant de Grupo P (`a8cb009b-98b1-4b87-ab9f-ecee5d1ebd14`) y el plan Periferia Urbana
   (ID `9Ta9UufHO0mDGpATZEM5aGUAGnOz`, tomado del enlace de Planner en Teams) ya vienen
   preconfigurados en la página; solo hace falta el Client ID del registro de Azure.

3. Cada persona inicia sesión con su cuenta de Microsoft la primera vez; después entra directo.

### 4. Actualización
Automática: la página consulta Planner en vivo al abrirse y se refresca cada 5 minutos
(botón **Actualizar ahora** para forzarla). No hay archivos que regenerar.

**Sin configuración todavía:** el botón **"Ver con datos de ejemplo"** muestra el dashboard completo
con datos ficticios para revisar el diseño antes de conectarlo.

## Corte desde Excel (sin conexión): `build_from_planner.py`

El dashboard trae **embebido un corte real del plan** (export de Excel de Planner), así que funciona
completo aunque no haya conexión configurada; la barra superior indica la fecha del corte. Para
refrescar el corte con un export nuevo:

1. En Planner (Teams) → el plan → **… → Exportar plan a Excel**.
2. `python3 build_from_planner.py "PLANNER_PERIFERIA_URBANA.xlsx"` (requiere `pip install openpyxl`).
3. Commit y push de `planner-kpi.html`.

La conexión en vivo (botón "Conectar con Microsoft") sigue disponible y, cuando está configurada,
tiene prioridad sobre el corte. Corte actual: **07/07/2026 · 501 tareas**.
