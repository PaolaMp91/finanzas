# Dashboard Financiero · Boulevard Sur Club Residencial

Dashboard interactivo **consolidado** (un solo archivo HTML, diseño corporativo) con el resumen
financiero del proyecto **Boulevard Sur Club Residencial** (Periferia Urbana · Mixco, Guatemala),
integrando **Fase 1 + Fase 2**.

## Cifras consolidadas del proyecto

- **704 apartamentos** (Fase 1: 216 · Fase 2: 488)
- **875 parqueos** netos (Fase 1: 321 · Fase 2: 594, descontando 40 de doble conteo)
- Fase 1 (Torres A·B·C, 72 aptos c/u) en obra; Fase 2 (Torres D·E·F) en evaluación con 4 escenarios.

## Cómo usarlo

Abra **`index.html`** en cualquier navegador (doble clic). No requiere instalación.
Necesita conexión a internet la primera vez para cargar la librería de gráficas (Chart.js por CDN).

## Qué muestra

| Sección | Contenido |
|---|---|
| **KPIs principales** | TIR de accionistas, TIR de proyecto, margen, utilidad proyectada, total ventas, inversión total, requerimiento de capital y años de ejecución. |
| **Selector de escenario** | Cambia entre los 4 escenarios de Fase 2 (E1–E4); todos los KPIs y la gráfica de TIR se actualizan. |
| **Unidades y parqueos** | 488 apartamentos (2D: 334 · 3D: 154), composición de torres (D/E/F + TP2) y parqueos por fase (F1: 321 · F2: 594 · total neto 875). |
| **Comparativo de escenarios** | Tabla y gráfica de ventas, utilidad, margen y TIR de los 4 escenarios. |
| **Resumen de costos** | Presupuesto de inversión Fase 2 (Q 255.4 M) por partida con porcentajes + tasa de banco (8%). |
| **Cronograma y capital** | Inicio de obra (Feb 2027), entrega (Feb 2030), crédito bancario, requerimiento de capital y comparativo bancario de Fase 1 (BI vs G&T, TIR socios 31.20%). |
| **Resumen de desembolsos** | Crédito autorizado, mecánica de desembolso, capital aportado y gastos financieros. |
| **Flujo de costos por mes** | Curva mensual proyectada vs. real con tabla **editable** (columna *Real* recalcula variación, % y gráfica). |

## Datos editables

Todos los datos están en el bloque `DATOS DEL PROYECTO` al inicio del `<script>` en `index.html`
(`ESCENARIOS`, `COSTOS`, etc.). Actualícelos con el archivo maestro cuando cambien las cifras.
La columna **Real** del flujo mensual se edita directamente con clic en la celda.

## Fuentes

Microsoft Teams › Periferia Urbana › Canal **Boulevard Sur** › carpeta **Flujo Financiero**:

- `PU PROJECT 94M Escenario BI 31-05-2026 12 nvls.xlsx` (archivo maestro)
- `2026-06-17 - Escenarios Financieros F2.pptx` (escenarios E1–E4, unidades, parqueos, cronograma)
- `2026-21-05 Presentación Financiera` (comparativo bancario BI vs G&T · Fase 1)
- `HANDOVER – BOULEVARD SUR CLUB RESIDENCIAL.pdf` (unidades de Fase 1: 216 aptos)
- Minuta de reunión financiera 11-06-2026

> Nota: el archivo maestro Excel (16 MB) no pudo convertirse a texto automáticamente; las cifras
> provienen de las presentaciones financieras derivadas de ese análisis. El calendario detallado de
> desembolsos y el flujo mensual real deben cargarse desde el archivo maestro.
