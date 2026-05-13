# CLAUDE.md — Consolidado SII / Dashboard Contable

## Descripción del proyecto

Dashboard contable para visualizar datos tributarios exportados del SII (Servicio de Impuestos Internos de Chile). El proyecto tiene dos capas:

1. **Excel enrichment** (`create_dashboard.py`) — lee el archivo SII, calcula KPIs y genera una hoja `Dashboard` con gráficos embebidos.
2. **Web dashboard** (`dashboard_sii.html`) — página HTML standalone para presentar los mismos datos con diseño elegante orientado a contadores.

El repositorio vive en: `https://github.com/AutomatizacionesBCH/RICHIPRUEBAS`

---

## Estructura de archivos

```
PRUEBAS RICHI/
├── Consolidado SII.xlsx      # Fuente de datos principal (NO editar manualmente)
│   ├── Dashboard             # Hoja auto-generada por create_dashboard.py
│   ├── Compras 2025          # 3,593 filas × 30 columnas — datos de compras
│   └── Ventas 2025           # 251 filas × 46 columnas — datos de ventas
├── create_dashboard.py       # Script que genera la hoja Dashboard en Excel
├── dashboard_sii.html        # Dashboard web standalone
├── Logo/
│   └── logo.png              # Logo geométrico verde (cubo arquitectónico)
└── .gitignore                # Excluye ~$*.xlsx y .firecrawl/
```

---

## Stack técnico

| Componente | Tecnología |
|---|---|
| Procesamiento Excel | Python 3 + `openpyxl` |
| Web dashboard | HTML/CSS/JS puro (un solo archivo) |
| Gráficos web | Chart.js 4.4.0 (CDN) |
| Tipografía web | Cormorant Garamond (display) + DM Mono (números) + Cinzel (headings) vía Google Fonts |
| Servidor local | `python -m http.server 8765` desde el directorio raíz |

---

## Dependencias Python

```bash
pip install openpyxl
```

No hay `requirements.txt`. La única librería externa es `openpyxl`.

---

## Cómo ejecutar

### Regenerar la hoja Dashboard en Excel

```bash
python create_dashboard.py
```

El script abre `Consolidado SII.xlsx`, elimina la hoja `Dashboard` si existe, la recrea con KPIs y tres gráficos (barras por mes, pie de tipo de compra, barras horizontales de proveedores), y guarda el archivo.

**Importante:** el archivo Excel debe estar **cerrado** antes de ejecutar el script. Si está abierto, openpyxl no puede sobreescribirlo.

### Levantar el servidor web

```bash
python -m http.server 8765
```

Luego abrir: `http://localhost:8765/dashboard_sii.html`

---

## Columnas relevantes de "Compras 2025"

| Índice (0-based) | Campo usado en el script |
|---|---|
| 0 | Empresa |
| 1 | Mes |
| 4 | Tipo de compra (`tc`) |
| 6 | Proveedor |
| 11 | Monto exento |
| 12 | Monto neto |
| 13 | IVA |
| 16 | Monto total |

---

## Diseño visual

### Excel Dashboard
- Fondo negro (`#000000`) en todas las celdas
- Texto en rosa (`#FF69B4`) y rosa fuerte (`#FF1493`)
- Bordes en `#FF1493`
- Cards KPI con fondo `#120010`
- Tres gráficos embebidos via openpyxl

### Web Dashboard (`dashboard_sii.html`)
- Paleta: verde oscuro `#070e07` + negro carbón, acentos en oro `#c9a84c`
- Grid de fondo con líneas doradas sutiles (`rgba(201,168,76,0.025)`)
- KPIs con animación count-up al cargar (easing cúbico, 1600ms)
- Barras de empresas y proveedores con reveal animado via `IntersectionObserver`
- Chart.js con tooltips estilizados en verde/dorado
- Totalmente standalone (un solo archivo HTML, sin build)

---

## Datos del dashboard (hardcoded en el HTML)

Los valores están embebidos directamente en `dashboard_sii.html`. Si cambian los datos del Excel, hay que actualizar el HTML manualmente o automatizar la extracción.

| KPI | Valor |
|---|---|
| Total documentos | 3,593 |
| Monto Neto | $301.94M CLP |
| IVA Recuperable | $53.81M CLP |
| Monto Total | $369.78M CLP |
| Monto Exento | $6.51M CLP |

**Empresas:** Puerto Querido SpA (81.3%), My Kit SpA (11.2%), Profesionales de la Salud SpA (5.5%), Los Dibujantes SpA (1.0%), VR Therapy SpA (1.0%)

---

## Convenciones de código

- `create_dashboard.py` usa rutas absolutas hardcoded (Windows). Si se mueve el proyecto, actualizar la ruta en las líneas `load_workbook(...)` y `wb.save(...)`.
- El HTML no tiene framework ni build step. Editar directamente.
- No hay tests automatizados.

---

## Git

Rama principal: `master`
Remote: `https://github.com/AutomatizacionesBCH/RICHIPRUEBAS.git`

```bash
git add <archivos>
git commit -m "descripción"
git push origin master
```
