# Modernizar tabs/botones/inputs (Tigo + WOM) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cambiar tabs (subrayado -> pill), y subir border-radius de botones/inputs de 8px a 10px en `informe_clima_app.html` (Tigo) y `informe_wom_app.html` (WOM), sin tocar layout, modales, toasts, cards ni tablas.

**Architecture:** Cambio puro CSS dentro del `<style>` inline de cada archivo HTML. Sin JS, sin build step. Cada archivo se edita y se verifica de forma independiente.

**Tech Stack:** HTML/CSS estatico servido por Express (`server.js`, ya corriendo en `http://localhost:3000`).

## Global Constraints

- No modificar modales (`.modal-*`, `.confirm-*`, `.cbtn`), toasts (`#toast`), cards de historial/papelera, ni tablas de mediciones — fuera de alcance segun spec `docs/superpowers/specs/2026-06-19-modernizar-componentes-design.md`.
- Cada app conserva su color de acento propio: `--blue` en Tigo, `--mag` en WOM. No unificar paleta.
- No hay test runner en este proyecto — verificacion es visual (curl + revision manual en navegador).

---

### Task 1: Tigo — tabs pill + botones + inputs

**Files:**
- Modify: `informe_clima_app.html:43-60` (bloque `/* ── TABS ── */`)
- Modify: `informe_clima_app.html:195` (`.btn-vaciar`)
- Modify: `informe_clima_app.html:211` (`.act-restore`)
- Modify: `informe_clima_app.html:213` (`.act-perm`)
- Modify: `informe_clima_app.html:220` (`.btn-nuevo-start`)
- Modify: `informe_clima_app.html:228` (`.btn-refresh`)
- Modify: `informe_clima_app.html:253` (`.act-btn`)
- Modify: `informe_clima_app.html:278` (`.btn`)
- Modify: `informe_clima_app.html:75` (regla generica de inputs)
- Modify: `informe_clima_app.html:86` (`.site-in`)
- Modify: `informe_clima_app.html:106` (`.sup-in`)
- Modify: `informe_clima_app.html:116` (`.tec-in`)
- Modify: `informe_clima_app.html:123` (`.sala-in`)
- Modify: `informe_clima_app.html:130` (`.marca-in`)
- Modify: `informe_clima_app.html:137` (`.tipo-in`)
- Modify: `informe_clima_app.html:193` (`.pap-toolbar input`)
- Modify: `informe_clima_app.html:226` (`.hist-search-wrap input`)

**Interfaces:**
- Produces: patron de pill-track (`.tabs-inner{background:#EDEEF2;border-radius:12px;padding:4px;width:fit-content;}` + `.tab{border-radius:9px;border:none;}` + `.tab.active{background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.12);}`) que Task 2 replica con `--mag` en vez de `--blue`.
- Consumes: nada de tasks anteriores.

- [ ] **Step 1: Reemplazar bloque de tabs (subrayado -> pill)**

En `informe_clima_app.html`, reemplazar exactamente:

```css
/* ── TABS ── */
.tabs{background:#fff;border-bottom:2px solid var(--border);overflow-x:auto;scrollbar-width:none;position:sticky;top:57px;z-index:150;box-shadow:0 2px 8px rgba(0,0,0,.04);}
.tabs::-webkit-scrollbar{display:none;}
.tabs-inner{max-width:var(--max-w);margin:0 auto;display:flex;}
.tab{flex:0 0 auto;padding:13px 17px;font-size:13px;font-weight:500;color:var(--muted);cursor:pointer;border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;transition:.18s;}
.tab:hover{color:var(--blue);background:#F0F6FE;}
.tab.active{color:var(--blue);border-bottom-color:var(--blue);font-weight:700;}
.tab.locked{opacity:.3;cursor:not-allowed;pointer-events:none;}
.tab .num{display:inline-flex;align-items:center;justify-content:center;min-width:18px;height:18px;background:#EBF4FE;border-radius:10px;font-size:9px;font-weight:700;margin-right:6px;padding:0 4px;color:var(--blue);}
.tab.active .num{background:var(--blue);color:#fff;}
.tab.nuevo{color:var(--green);font-weight:600;}
.tab.nuevo .num{background:#D4F5E7;color:var(--green);}
.tab.nuevo.active{color:var(--green);border-bottom-color:var(--green);font-weight:700;}
.tab.nuevo.active .num{background:var(--green);color:#fff;}
.tab.papelera{color:#94a3b8;}
.tab.papelera .num{background:#f8fafc;color:#94a3b8;}
.tab.papelera.active{color:var(--red);border-bottom-color:var(--red);}
.tab.papelera.active .num{background:var(--red);color:#fff;}
```

por:

```css
/* ── TABS ── */
.tabs{background:#fff;border-bottom:2px solid var(--border);overflow-x:auto;scrollbar-width:none;position:sticky;top:57px;z-index:150;box-shadow:0 2px 8px rgba(0,0,0,.04);padding:8px 0;}
.tabs::-webkit-scrollbar{display:none;}
.tabs-inner{max-width:var(--max-w);margin:0 auto;display:flex;gap:4px;background:#EDEEF2;border-radius:12px;padding:4px;width:fit-content;}
.tab{flex:0 0 auto;padding:9px 16px;font-size:13px;font-weight:500;color:var(--muted);cursor:pointer;border:none;border-radius:9px;white-space:nowrap;transition:.18s;}
.tab:hover{color:var(--blue);}
.tab.active{color:var(--blue);background:#fff;font-weight:700;box-shadow:0 1px 3px rgba(0,0,0,.12);}
.tab.locked{opacity:.3;cursor:not-allowed;pointer-events:none;}
.tab .num{display:inline-flex;align-items:center;justify-content:center;min-width:18px;height:18px;background:#EBF4FE;border-radius:10px;font-size:9px;font-weight:700;margin-right:6px;padding:0 4px;color:var(--blue);}
.tab.active .num{background:var(--blue);color:#fff;}
.tab.nuevo{color:var(--green);font-weight:600;}
.tab.nuevo .num{background:#D4F5E7;color:var(--green);}
.tab.nuevo.active{color:var(--green);background:#fff;font-weight:700;}
.tab.nuevo.active .num{background:var(--green);color:#fff;}
.tab.papelera{color:#94a3b8;}
.tab.papelera .num{background:#f8fafc;color:#94a3b8;}
.tab.papelera.active{color:var(--red);background:#fff;}
.tab.papelera.active .num{background:var(--red);color:#fff;}
```

- [ ] **Step 2: Bump border-radius de botones (8px -> 10px), uno por uno**

| Selector | Linea aprox | Cambio |
|---|---|---|
| `.btn-vaciar` | 195 | `border-radius:8px` -> `border-radius:10px` |
| `.act-restore` | 211 | `border-radius:8px` -> `border-radius:10px` |
| `.act-perm` | 213 | `border-radius:8px` -> `border-radius:10px` |
| `.btn-nuevo-start` | 220 | `border-radius:8px` -> `border-radius:10px` |
| `.btn-refresh` | 228 | `border-radius:8px` -> `border-radius:10px` |
| `.act-btn` | 253 | `border-radius:8px` -> `border-radius:10px` |
| `.btn` | 278 | `border-radius:8px` -> `border-radius:10px` |

Cada fila: usar Edit con `old_string` = linea completa de la regla CSS tal como esta en el archivo (incluye el selector completo para que el match sea unico) y `new_string` = misma linea con `8px` cambiado a `10px` en la propiedad `border-radius`. No tocar ningun otro `border-radius:8px` que no este en esta tabla (hay otros en inputs, se tratan en el Step 3).

- [ ] **Step 3: Bump border-radius de inputs (8px -> 10px), uno por uno**

| Selector | Linea aprox |
|---|---|
| Regla generica `input[type="text"],input[type="date"],input[type="number"],input[type="email"],textarea` | 75 |
| `.site-in` | 86 |
| `.sup-in` | 106 |
| `.tec-in` | 116 |
| `.sala-in` | 123 |
| `.marca-in` | 130 |
| `.tipo-in` | 137 |
| `.pap-toolbar input` | 193 |
| `.hist-search-wrap input` | 226 |

Mismo metodo: Edit con `old_string` = linea completa de cada regla, cambiar solo `border-radius:8px` a `border-radius:10px`. No tocar `.dd`, `.sup-dd`, `.tec-dd`, `.sala-dd`, `.marca-dd`, `.tipo-dd` (ya estan en `10px`, quedan igual).

- [ ] **Step 4: Verificar que el servidor sirve el archivo actualizado**

Run: `curl -s -o /dev/null -w "%{http_code}" http://localhost:3000/tigo`
Expected: `200`

Run: `curl -s http://localhost:3000/tigo | grep -c "background:#EDEEF2"`
Expected: `1` (confirma que el nuevo bloque de tabs esta presente)

- [ ] **Step 5: Revision visual manual**

Abrir `http://localhost:3000/tigo` en navegador. Confirmar:
- Tabs se ven como pill (fondo gris, activa blanca con sombra), click cambia tab activo, scroll horizontal sigue funcionando con varias tabs.
- Tab "Nuevo" activa se ve verde, tab "Papelera" activa se ve roja (si hay items en papelera).
- Botones (Generar, Vaciar papelera, Restaurar, etc) se ven con esquina mas redonda, sin romper texto/iconos.
- Inputs (campos de texto, dropdowns de sitio/supervisor/tecnico/sala/marca/tipo) se ven con esquina mas redonda, focus-glow sigue funcionando al hacer click.

- [ ] **Step 6: Commit**

```bash
git add informe_clima_app.html
git commit -m "$(cat <<'EOF'
Modernizar tabs/botones/inputs en Tigo (estilo soft-rounded)

Tabs pasan de subrayado a pill-track; botones e inputs suben
border-radius de 8px a 10px. Sin cambios de layout ni en
modales/toasts/cards/tablas.
EOF
)"
```

---

### Task 2: WOM — tabs pill + botones + inputs

**Files:**
- Modify: `informe_wom_app.html:41-53` (bloque `/* ── TABS ── */`)
- Modify: `informe_wom_app.html:88` (`.btn-add-tec`)
- Modify: `informe_wom_app.html:118` (`.btn`)
- Modify: `informe_wom_app.html:139` (`.btn-dl`)
- Modify: `informe_wom_app.html:141` (`.btn-del`)
- Modify: `informe_wom_app.html:67` (`.field input,.field select,.field textarea`)

**Interfaces:**
- Consumes: patron de pill-track de Task 1 (`.tabs-inner`/`.tab`/`.tab.active`), adaptado a `--mag` en vez de `--blue`. WOM no tiene variantes `.tab.nuevo`/`.tab.papelera`.
- Produces: nada (ultimo task del plan).

- [ ] **Step 1: Reemplazar bloque de tabs (subrayado -> pill)**

En `informe_wom_app.html`, reemplazar exactamente:

```css
/* ── TABS ── */
.tabs{background:#fff;border-bottom:2px solid var(--border);overflow-x:auto;scrollbar-width:none;position:sticky;top:57px;z-index:150;box-shadow:0 2px 8px rgba(0,0,0,.04);}
.tabs::-webkit-scrollbar{display:none;}
.tabs-inner{max-width:var(--max-w);margin:0 auto;display:flex;}
.tab{flex:0 0 auto;padding:13px 17px;font-size:13px;font-weight:500;color:var(--muted);cursor:pointer;border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;transition:.18s;}
.tab:hover{color:var(--mag);background:#F5F0FC;}
.tab.active{color:var(--mag);border-bottom-color:var(--mag);font-weight:700;}
.tab.locked{opacity:.3;cursor:not-allowed;pointer-events:none;}
.tab .num{display:inline-flex;align-items:center;justify-content:center;min-width:18px;height:18px;background:#EDE7F6;border-radius:10px;font-size:10px;font-weight:700;margin-right:6px;padding:0 4px;color:var(--mag);line-height:1;vertical-align:middle;flex-shrink:0;}
.tab.active .num{background:var(--mag);color:#fff;}
.tab .num.num-star{font-size:11px;padding-bottom:1px;}
.tab .num.num-win{font-size:0;padding:0;}
.tab .num.num-win svg{display:block;}
```

por:

```css
/* ── TABS ── */
.tabs{background:#fff;border-bottom:2px solid var(--border);overflow-x:auto;scrollbar-width:none;position:sticky;top:57px;z-index:150;box-shadow:0 2px 8px rgba(0,0,0,.04);padding:8px 0;}
.tabs::-webkit-scrollbar{display:none;}
.tabs-inner{max-width:var(--max-w);margin:0 auto;display:flex;gap:4px;background:#EDEEF2;border-radius:12px;padding:4px;width:fit-content;}
.tab{flex:0 0 auto;padding:9px 16px;font-size:13px;font-weight:500;color:var(--muted);cursor:pointer;border:none;border-radius:9px;white-space:nowrap;transition:.18s;}
.tab:hover{color:var(--mag);}
.tab.active{color:var(--mag);background:#fff;font-weight:700;box-shadow:0 1px 3px rgba(0,0,0,.12);}
.tab.locked{opacity:.3;cursor:not-allowed;pointer-events:none;}
.tab .num{display:inline-flex;align-items:center;justify-content:center;min-width:18px;height:18px;background:#EDE7F6;border-radius:10px;font-size:10px;font-weight:700;margin-right:6px;padding:0 4px;color:var(--mag);line-height:1;vertical-align:middle;flex-shrink:0;}
.tab.active .num{background:var(--mag);color:#fff;}
.tab .num.num-star{font-size:11px;padding-bottom:1px;}
.tab .num.num-win{font-size:0;padding:0;}
.tab .num.num-win svg{display:block;}
```

- [ ] **Step 2: Bump border-radius de botones (8px -> 10px), uno por uno**

| Selector | Linea aprox |
|---|---|
| `.btn-add-tec` | 88 |
| `.btn` | 118 |
| `.btn-dl` | 139 |
| `.btn-del` | 141 |

Mismo metodo que Task 1 Step 2: Edit linea completa, solo cambiar `border-radius:8px` a `border-radius:10px`. No tocar `.btn-icon` (6px, fuera de alcance, es boton mini de icono).

- [ ] **Step 3: Bump border-radius de inputs (8px -> 10px)**

Editar `informe_wom_app.html:67`:

```css
.field input,.field select,.field textarea{width:100%;border:1.5px solid var(--border);border-radius:8px;padding:10px 12px;font-size:14px;font-family:inherit;color:var(--text);background:#fff;transition:border-color .18s,box-shadow .18s;outline:none;}
```

por:

```css
.field input,.field select,.field textarea{width:100%;border:1.5px solid var(--border);border-radius:10px;padding:10px 12px;font-size:14px;font-family:inherit;color:var(--text);background:#fff;transition:border-color .18s,box-shadow .18s;outline:none;}
```

No tocar `.foto-item .cap input` (6px, caption de foto, fuera de alcance).

- [ ] **Step 4: Verificar que el servidor sirve el archivo actualizado**

Run: `curl -s -o /dev/null -w "%{http_code}" http://localhost:3000/wom`
Expected: `200`

Run: `curl -s http://localhost:3000/wom | grep -c "background:#EDEEF2"`
Expected: `1`

- [ ] **Step 5: Revision visual manual**

Abrir `http://localhost:3000/wom` en navegador. Confirmar:
- Tabs se ven como pill (fondo gris, activa blanca con sombra, texto/numero en magenta), click cambia tab activo, scroll horizontal sigue funcionando.
- Botones (Agregar tecnico, Descargar, Eliminar) se ven con esquina mas redonda.
- Inputs/selects del formulario se ven con esquina mas redonda, focus-glow magenta sigue funcionando.

- [ ] **Step 6: Commit**

```bash
git add informe_wom_app.html
git commit -m "$(cat <<'EOF'
Modernizar tabs/botones/inputs en WOM (estilo soft-rounded)

Mismo cambio que Tigo: tabs pasan a pill-track, botones e inputs
suben border-radius de 8px a 10px, manteniendo acento magenta propio
de WOM.
EOF
)"
```
