# Dashboard PMO conectado a Teams — puesta en marcha

Este dashboard (`dashboard.html`) se **actualiza solo**: toma de la carpeta de
Teams el archivo de *Seguimiento Semanal* con la **fecha más cercana a hoy**, lee
el Excel y regenera la página. Tú pegas **un solo link** en Teams y siempre
muestra la última versión.

```
Carpeta origen (Teams / SharePoint):
TECNICO-BOULEVARSUR › PLANIFICACION › 01. PLANNER › FASE 2 › 02. Seguimiento Semanal
```

## Cómo funciona

```
┌ Teams / SharePoint ┐     ┌ GitHub Actions (diario) ┐     ┌ GitHub Pages ┐
│ 02. Seguimiento     │ →   │ refresh_from_teams.py    │ →  │ dashboard.html│ →  Link en Teams
│ Semanal (xlsx)      │     │  · elige fecha + cercana │     │ (se refresca) │
└─────────────────────┘     │  · lee Excel             │     └───────────────┘
                            │  · genera el HTML        │
                            └──────────────────────────┘
```

1. **GitHub Actions** corre todos los días (y cuando lo dispares a mano).
2. `refresh_from_teams.py` se conecta a SharePoint con **Microsoft Graph**
   (solo lectura), descarga el `.xlsx` con la fecha más cercana a hoy y llama a
   `build_dashboard.py`.
3. El HTML se publica en **GitHub Pages** con una URL fija. Esa URL es la que
   pegas en Teams.

## Lo único que hay que configurar una vez (esto normalmente lo hace IT)

Para que el robot lea la carpeta **sin tu contraseña**, se registra una
*aplicación* en Azure AD con permiso de **solo lectura**.

### 1) Registrar la aplicación en Azure AD
- Portal de Azure → **Microsoft Entra ID / Azure AD** → **App registrations** →
  **New registration**. Nombre: `PMO Dashboard Boulevard Sur`. Registrar.
- Copiar **Directory (tenant) ID** y **Application (client) ID**.

### 2) Dar permiso de solo lectura
- En la app → **API permissions** → **Add a permission** → **Microsoft Graph**
  → **Application permissions** → agregar **`Sites.Read.All`**
  (o `Files.Read.All`).
- Clic en **Grant admin consent** (requiere un administrador del tenant).

### 3) Crear un secreto
- En la app → **Certificates & secrets** → **New client secret** → copiar el
  **Value** (solo se muestra una vez).

### 4) Guardar los 3 datos como Secrets del repositorio de GitHub
En `github.com/paolamp91/finanzas` → **Settings** → **Secrets and variables** →
**Actions** → **New repository secret**, crear:

| Secret | Valor |
|---|---|
| `AZURE_TENANT_ID` | Directory (tenant) ID |
| `AZURE_CLIENT_ID` | Application (client) ID |
| `AZURE_CLIENT_SECRET` | el Value del secreto |

### 5) Activar GitHub Pages
- Repo → **Settings** → **Pages** → **Source: GitHub Actions**.
- (Si el repositorio es privado, GitHub Pages requiere plan **Pro/Team**. Si no,
  se puede hacer el repo público o publicar el HTML en otro hosting.)

### 6) Ejecutar la primera vez
- Repo → pestaña **Actions** → workflow **“Actualizar dashboard PMO desde Teams”**
  → **Run workflow**. Al terminar, la URL de Pages queda lista:

  ```
  https://paolamp91.github.io/finanzas/dashboard.html
  ```

## Pegar el link en Teams
- **Como pestaña:** en el canal → **+** → **Website** → pega la URL. Cada vez que
  alguien la abra verá la última versión (la página además se auto-refresca sola).
- **Como mensaje:** simplemente pega la URL en el chat/canal.

## Cambiar cada cuánto se actualiza
En `.github/workflows/refresh-dashboard.yml`, la línea `cron: '0 12 * * *'`
(diario 06:00 Guatemala). Ejemplos:
- Lunes y jueves 7am GT: `cron: '0 13 * * 1,4'`
- Cada 6 horas: `cron: '0 */6 * * *'`

## Probar el generador sin credenciales
```bash
pip install openpyxl
python3 build_dashboard.py "13 07 2026 DASHBOARD PLANIFICACION.xlsx"   # con un Excel descargado
python3 build_dashboard.py                                             # vista previa con datos de semilla
```

## Notas sobre los datos
El Excel es un export de **MS Project**. El dashboard usa las columnas
`Name, Outline_Level, Percent_Complete, Start, Finish, Duration, Resource_Names`.
El **% de avance por fase** se toma del valor acumulado que MS Project calcula en
las filas de resumen (nivel 1 y 2). Si una semana el export no trae la columna
`Percent_Complete`, esas celdas aparecerán como “sincroniza” hasta el próximo
archivo que sí la incluya.
