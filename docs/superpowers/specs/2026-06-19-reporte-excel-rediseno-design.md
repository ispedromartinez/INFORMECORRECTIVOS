# Reporte Excel: rediseño (periodo en vez de informes, estilo, checklist vertical)

## Alcance

Reemplaza el modal "Reporte" recien implementado (`docs/superpowers/specs/2026-06-19-reporte-excel-historial-design.md`)
en `informe_clima_app.html` (Tigo), `informe_wom_app.html` (WOM) y los
endpoints `/reporte`/`/reporte-wom` en `server.js`. No toca el resto del
sistema.

## Que cambia respecto al modal actual

1. **Se quita el checklist de informes individuales.** Ya no se eligen
   informes uno por uno.
2. **Se agrega seleccion de periodo**: tabs "Todos / Mes / Semana / Fecha"
   (mismo patron visual pill-track que las tabs de navegacion principal).
   "Todos" viene seleccionado por defecto al abrir el modal.
   - **Mes**: `<input type="month">`.
   - **Semana**: `<input type="date">` — el usuario elige cualquier dia
     de la semana que quiere; el servidor calcula el rango lunes-domingo
     que contiene esa fecha.
   - **Fecha**: dos `<input type="date">` (Desde/Hasta), mismo patron que
     los campos "Fecha Inicio"/"Fecha Término" que Tigo ya usa en la
     pestaña General.
3. **El filtro de periodo usa `fechaCreacion`** (timestamp ISO real de
   cuando se generó el informe en el sistema), no el campo `fecha`
   (texto libre que escribe el técnico, a veces un rango "19/06 - 25/06").
   No se persiste ningún campo nuevo — los rangos de mes/semana/fecha se
   calculan al vuelo contra `fechaCreacion` en el momento de generar el
   reporte. Esto resuelve la idea de "registro interno por semana" sin
   agregar ningún campo ni UI nueva visible: es pura aritmética de
   fechas sobre un dato que ya existe.
4. **El checklist de columnas pasa a vertical** (una por línea, no chips
   en fila).
5. **Estilo del modal**: se reusa el patrón de tarjeta que ya existe en
   el resto del sistema (`.card-title` con barra de color de 4px +
   título en mayúsculas) para las dos secciones del modal ("Periodo" y
   "Columnas"), en vez de los bloques de texto plano que tiene hoy.
   Tipografía: la que ya define el `<body>` (Poppins), sin agregar
   fuentes nuevas.

## Backend

**Contrato nuevo** para `POST /reporte` y `POST /reporte-wom`:

```json
{
  "columns": ["fecha", "nombreSitio"],
  "periodo": { "tipo": "todos" }
}
```

`periodo.tipo` es uno de: `"todos"`, `"mes"`, `"semana"`, `"fecha"`.

- `"todos"`: sin parámetros extra, incluye todo el historial.
- `"mes"`: `{ "tipo":"mes", "mes":"2026-06" }` (formato `input type="month"`).
- `"semana"`: `{ "tipo":"semana", "fecha":"2026-06-19" }` — el servidor
  calcula el lunes 00:00:00 y el domingo 23:59:59 de la semana que
  contiene esa fecha (semana empieza lunes).
- `"fecha"`: `{ "tipo":"fecha", "desde":"2026-06-01", "hasta":"2026-06-19" }`.

El servidor obtiene todos los registros (`dbClimaList(null)` /
`dbWomList()`), calcula el rango `[inicio, fin]` según `periodo`, y se
queda con los registros cuyo `fechaCreacion` cae en ese rango (para
`"todos"` no filtra nada). Ya no recibe ni usa `ids`. Si `periodo.tipo`
no es uno de los 4 valores válidos, o si `mes`/`fecha`/`desde`/`hasta`
faltan para el tipo correspondiente, responde 400.

Las columnas exportables y sus headers en español **no cambian** respecto
al spec anterior (mismo `REPORTE_COLUMNAS_CLIMA`/`REPORTE_COLUMNAS_WOM`).

## Frontend

**Modal** (misma estructura `.modal-overlay`/`.modal-box` que hoy, con
contenido reemplazado):

- Sección "Periodo" envuelta en una tarjeta con barra de color (mismo
  acento de cada app: `var(--blue)` en Tigo, `var(--mag)` en WOM — igual
  que usa `.card-title::before` en el resto del sistema, sin variar color
  por sección) + título "PERIODO": tabs Todos/Mes/Semana/Fecha; debajo de
  las tabs, el input correspondiente al tipo activo (oculto para "Todos",
  `<input type="month">` para Mes, `<input type="date">` para Semana, dos
  `<input type="date">` para Fecha).
- Sección "Columnas" en su propia tarjeta, mismo estilo de barra +
  título ("COLUMNAS"), mismo acento (no se introduce un segundo color):
  checklist vertical (una columna por línea), todas marcadas por
  defecto.
- Botón "Generar Excel" deshabilitado si no hay ninguna columna marcada,
  o si el tipo de periodo activo requiere un input que está vacío (mes
  sin elegir, semana sin fecha, o fecha sin desde/hasta).

**JS**: `generarReporte()` ya no junta `ids` de checkboxes; arma el
objeto `periodo` según la tab activa y lo envía junto con `columns` al
mismo patrón fetch+blob+download que ya existe.

## Fuera de alcance

Seguir igual que el spec anterior: no se agregan columnas nuevas
(LPU/tickets/resumen/observaciones), no hay exportación en otros
formatos, no hay reportes programados. Tampoco se valida que
"semana"/"mes" tengan datos antes de generar — si no hay informes en el
rango, el Excel sale con solo el header (mismo comportamiento ya
verificado en el backend actual).

## Testing

Sin test runner. Verificación manual: generar informes con fechas de
creación en distintas semanas/meses (no se puede forzar `fechaCreacion`
desde la UI, así que la prueba real es generar 2-3 informes hoy y
confirmar que "Semana"/"Mes"/"Fecha" de hoy los incluye, y un rango de
fecha que no cubre hoy los excluye). Confirmar que "Todos" sigue
trayendo todo. Confirmar visualmente que el checklist de columnas se ve
vertical y que las secciones tienen la barra de color + título como el
resto del sistema.
