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

Bump de `border-radius` de 8px a 10px en los botones rectangulares de
ambos archivos (`.btn`, `.btn-vaciar`, `.btn-nuevo-start`,
`.btn-refresh`, `.act-restore`, `.act-perm`, `.act-btn` en Tigo;
`.btn-add-tec`, `.btn`, `.btn-dl`, `.btn-del` en WOM). Excluidos:
`.btn-gen` (ya usa `var(--r)`=12px), `.cbtn`/`.modal-close` (viven en
modales, fuera de alcance), `.rm-btn`/`.rst-btn`/`.rot-btn` (circulos,
no aplica radio), `.back-btn`/`.btn-icon` (chrome menor, no formulario).
Sombras (`box-shadow`) existentes se mantienen sin cambio — ya estan en
linea con el estilo "soft".

### Inputs

Bump de `border-radius` de 8px a 10px en los campos de texto/fecha/
numero/email/textarea de Tigo (regla generica, `.site-in`, `.sup-in`,
`.tec-in`, `.sala-in`, `.marca-in`, `.tipo-in`, `.pap-toolbar input`,
`.hist-search-wrap input`) y en `.field input/select/textarea` de WOM.
Los menus desplegables (`.dd`, `.sup-dd`, `.tec-dd`, `.sala-dd`,
`.marca-dd`, `.tipo-dd`) ya estan en 10px, quedan igual. Borde (1.5px) y
focus-glow (`box-shadow` con color del accent) se mantienen igual.

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
