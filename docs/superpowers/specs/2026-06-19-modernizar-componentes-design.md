# Modernizar componentes UI (Tigo + WOM)

## Alcance

`informe_clima_app.html` (Tigo) y `informe_wom_app.html` (WOM). No incluye
`selector.html`, modales, toasts, cards ni tablas.

## Direccion visual

"Soft rounded": esquinas mas redondeadas, sombras suaves, tabs tipo pill.
Cada app mantiene su color de acento existente (`--blue` en Tigo, `--mag`
en WOM) — no se unifica paleta.

## Componentes

### Tabs

Reemplazar indicador de subrayado (`border-bottom` en `.tab.active`) por
estilo pill-track:

- `.tabs-inner` gana fondo gris claro (`#eceef2`), `border-radius:12px`,
  padding interno ~4px, tabs en fila con gap pequeno.
- `.tab` pasa a `border-radius:9px`, sin borde inferior.
- `.tab.active` fondo blanco + `box-shadow` suave + color de texto = accent
  de la app.
- Estados especiales se conservan con la misma logica de color pero forma
  pill: en Tigo, `.tab.nuevo.active` queda verde, `.tab.papelera.active`
  rojo; en WOM no hay esos estados.
- Scroll horizontal (`overflow-x:auto`) y badges `.num` se mantienen
  igual, solo ajustando el radio del badge si choca visualmente con la
  pill.

### Botones

Bump de `border-radius` de 8px a 10px en todos los botones de ambos
archivos (`.btn`, `.btn-primary`, `.btn-secondary`, `.btn-gen`,
`.btn-refresh`, `.btn-vaciar`, `.btn-add-tec`, `.btn-dl`, `.btn-del`,
`.act-*`, `.cbtn`, `.rm-btn`/`.rst-btn`/`.rot-btn` si aplica visualmente).
Sombras (`box-shadow`) existentes se mantienen sin cambio — ya estan en
linea con el estilo "soft".

### Inputs

Bump de `border-radius` de 8px a 10px en todos los campos de texto,
fecha, numero, email, textarea, select y los dropdowns custom de Tigo
(`.site-in`, `.sup-in`, `.tec-in`, `.sala-in`, `.marca-in`, `.tipo-in` y
sus `.dd`/`*-dd` desplegables) y de WOM (`.field input/select/textarea`).
Borde (1.5px) y focus-glow (`box-shadow` con color del accent) se
mantienen igual.

## Fuera de alcance

Modales, confirm-box, toasts, cards de historial/papelera, tablas de
mediciones, paleta de colores, layout/estructura general. Nada de esto se
toca en este cambio.

## Testing

Cambio es puramente CSS (`border-radius` + reestructura de `.tabs`/`.tab`).
Verificacion manual en navegador: abrir `/tigo` y `/wom`, confirmar tabs
pill funcionan (click cambia activo, scroll horizontal ok, estados
nuevo/papelera en Tigo se ven correctos), botones e inputs con radio
nuevo sin romper layout.
