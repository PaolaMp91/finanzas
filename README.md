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

## Secciones del dashboard
Resumen consolidado real · Comparativa estratégica A vs B con recomendación · Gráficas de TIR y utilidad ·
Unidades por fase · Presupuesto de inversión F2 (con los cambios) · Cronograma y comparativo bancario BI/G&T ·
Flujo de costos de construcción mensual proyectado vs. real (editable, por escenario).

## Fuente
Microsoft Teams › Periferia Urbana › Canal **Boulevard Sur** › carpeta **Flujo Financiero** →
`PU PROJECT 94M Escenario BI 31-05-2026 12 nvls.xlsx` (hojas RESUMEN GENERAL, ER F2, SUPUESTOS F2, INDICES, INCREMENTO F2).
