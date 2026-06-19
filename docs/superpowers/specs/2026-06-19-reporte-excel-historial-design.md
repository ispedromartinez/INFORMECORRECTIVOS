# Exportar Excel desde Historial (Tigo + WOM)

## Alcance

`informe_clima_app.html` (Tigo), `informe_wom_app.html` (WOM), `server.js`.
Agrega un boton "Reporte" en la pestana Historial de cada app que abre un
modal de seleccion (columnas + informes) y genera un archivo `.xlsx`
descargable. Cada app tiene su propio boton/modal/endpoint — no comparten
datos entre si.

## Datos exportables

Solo campos que **ya se guardan hoy** en el historial (sin agregar
captura de datos nuevos):

**Tigo** (`/registro`): `fecha`, `codInforme`, `nombreSitio`,
`codigoSitio`, `tecnico`, `supervisor`, `numOT`, `photoCount`.

Headers en Excel: Fecha, Código Informe, Sitio, Código Sitio, Técnico,
Supervisor, OT, Cant. Fotos.

**WOM** (`/registro-wom`): `fechaInicio`, `ticket`, `codInterno`,
`instalacion`, `tipoActividad`, `tecnicos`, `photoCount`.

Headers en Excel: Fecha Inicio, Ticket, Código Interno, Instalación, Tipo
Actividad, Técnicos, Cant. Fotos.

LPU, numeros de ticket individuales (Inc/TE/TI/RED), resumen y
observaciones NO estan disponibles como columnas porque no se persisten
en el historial hoy (solo van al .docx) — fuera de alcance de este
cambio.

## UI

**Boton:** "📊 Reporte" junto al boton de refresh (`.btn-refresh`) en
`.hist-search-wrap`, mismo estilo `.btn-refresh`.

**Modal** (reusa estructura `.modal-overlay`/`.modal-box` ya existente en
ambas apps):
- Titulo: "Generar reporte Excel"
- Seccion 1 "Columnas": checkbox por cada campo exportable (lista de
  arriba), todos marcados por defecto.
- Seccion 2 "Informes a incluir": checkbox por cada informe actualmente
  visible en Historial (respeta el filtro de busqueda activo al momento
  de abrir el modal — usa el array ya cargado en memoria, no vuelve a
  pedir al servidor), todos marcados por defecto. Cada item muestra
  sitio/ticket + fecha para identificarlo.
- Boton "Generar Excel": deshabilitado si no hay columnas o informes
  marcados. Dispara la descarga y cierra el modal.
- Boton cerrar (X), igual que otros modales.

## Backend

**Tigo:** `POST /reporte` — body `{ columns: string[], ids: string[] }`.
Cae a `dbClimaList()`/lookup local segun ids, filtra por los ids
recibidos, construye una hoja de Excel con `xlsx` (SheetJS): primera fila
= headers en español de las columnas pedidas (mismo orden), filas
siguientes = valores correspondientes por informe. Devuelve el `.xlsx`
como buffer binario con
`Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
y `Content-Disposition: attachment`. Si `columns` o `ids` viene vacio,
responde 400.

**WOM:** `POST /reporte-wom`, mismo patron contra `dbWomList()`.

**Dependencia:** `xlsx` ya esta en `devDependencies` de `package.json`;
se mueve a `dependencies` porque ahora lo usa el servidor en runtime.

## Frontend (fetch + descarga)

Mismo patron que `generateDocx()` ya usa para `/generar`: `fetch` POST
con JSON body, `await resp.blob()`, `URL.createObjectURL(blob)`, link
temporal con `download="reporte.xlsx"` y click programatico. Sin
recargar la pagina ni depender de que el servidor setee el nombre de
archivo (se fija client-side, ej. `reporte-tigo-2026-06-19.xlsx`).

## Fuera de alcance

Agregar campos nuevos al historial (LPU, tickets, resumen,
observaciones), exportar fotos o el contenido del Word, programar
reportes automaticos, exportar en otros formatos (CSV/PDF).

## Testing

Sin test runner en el proyecto. Verificacion manual:
- Generar 2-3 informes de prueba (Tigo y WOM), abrir modal Reporte,
  confirmar que la lista de informes coincide con Historial.
- Desmarcar una columna y un informe, generar, abrir el `.xlsx`
  resultante y confirmar que excluye exactamente lo desmarcado.
- Probar con 0 columnas o 0 informes marcados: boton "Generar Excel"
  debe estar deshabilitado.
