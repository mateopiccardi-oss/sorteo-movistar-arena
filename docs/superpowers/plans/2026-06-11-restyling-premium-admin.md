# Restyling premium app admin — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Terminación premium del estilo actual (dirección A): tokens más generosos, profundidad real, badges pill, avatares con iniciales, focus rings — sin cambiar identidad, layout ni funcionalidad. Limpieza: eliminar `renderTasa()` (función muerta).

**Architecture:** Todo en `index.html`: edits puntuales al bloque `<style>` (líneas 15-555) y a 3 templates JS. Sin dependencias nuevas.

**Tech Stack:** HTML/CSS/JS vanilla.

**Spec:** `docs/superpowers/specs/2026-06-11-restyling-premium-admin-design.md`

---

### Task 1: Tokens y componentes base (CSS)

**Files:** Modify: `index.html` bloque `<style>`

- [ ] **Step 1:** En `:root` (línea 25) reemplazar `--r:8px;--rsm:5px;` por:

```css
  --r:12px;--rsm:9px;
  --grad-card:linear-gradient(160deg,rgba(26,40,64,0.95),rgba(17,26,41,0.95));
  --shadow-card:0 10px 26px rgba(0,0,0,0.32),inset 0 1px 0 rgba(255,255,255,0.06);
  --glow-ice:0 6px 20px rgba(0,212,255,0.32);
```

- [ ] **Step 2:** `.card` (línea 76): `background:rgba(22,32,48,0.9)` → `background:var(--grad-card)`, `border:1px solid rgba(255,255,255,0.09)` → `0.10`, y agregar `box-shadow:var(--shadow-card)`.

- [ ] **Step 3:** `.stat` (línea 85): mismos tres cambios que `.card`.

- [ ] **Step 4:** `.bp` (línea 93): `background:var(--ice)` → `background:linear-gradient(135deg,var(--ice2),var(--ice))`, `box-shadow:0 0 18px rgba(0,212,255,0.2)` → `box-shadow:var(--glow-ice)`. En `.bp:hover` (línea 395 bloque animaciones): shadow → `0 8px 26px rgba(0,212,255,0.42)`. En `.bp:hover` (línea 94): quitar `background:var(--ice2)` (el gradiente queda fijo, el lift ya da feedback) → reemplazar por `filter:brightness(1.08)`.

- [ ] **Step 5:** `.bo` (línea 95): agregar `background:rgba(255,255,255,0.03)` (reemplaza `background:transparent`).

- [ ] **Step 6:** `.field input:focus,...` (línea 111): agregar `box-shadow:0 0 0 3px rgba(0,212,255,0.10)`.

- [ ] **Step 7:** `.ph-t` (línea 72): `font-size:30px` → `32px`.

- [ ] **Step 8:** Agregar al final del CSS (antes de `</style>`):

```css
/* ── SCROLLBAR PREMIUM ── */
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.10);border-radius:99px;border:2px solid var(--bg)}
::-webkit-scrollbar-thumb:hover{background:rgba(255,255,255,0.18)}
/* ── AVATAR INICIALES ── */
.avi{width:26px;height:26px;border-radius:50%;background:linear-gradient(135deg,rgba(0,212,255,0.25),rgba(0,212,255,0.08));border:1px solid rgba(0,212,255,0.3);display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;color:var(--ice2);flex-shrink:0;letter-spacing:.5px}
.avi.gray{background:linear-gradient(135deg,rgba(255,255,255,0.10),rgba(255,255,255,0.04));border-color:rgba(255,255,255,0.14);color:var(--gray2)}
```

- [ ] **Step 9:** Commit: `git add index.html; git commit -m "style(admin): tokens premium, profundidad en cards/stats, boton primario con gradiente"`

---

### Task 2: Topbar, wizard, badges pill, toast, modal

**Files:** Modify: `index.html` bloque `<style>`

- [ ] **Step 1:** `.tb` (línea 35): `background:rgba(15,21,32,0.92)` → `rgba(13,19,30,0.85)`.

- [ ] **Step 2:** `.ws.on` (línea 55): agregar `background:linear-gradient(180deg,transparent 60%,rgba(0,212,255,0.07))`. `.ws.on .ws-n` (línea 58): `background:var(--ice)` → `background:linear-gradient(135deg,var(--ice2),var(--ice))` y agregar `box-shadow:0 0 14px rgba(0,212,255,0.45)`.

- [ ] **Step 3:** Badges a pill: `.bic` (línea 118) `border-radius:4px` → `99px`; `.ev-badge` (línea 259) `border-radius:3px` → `99px`; `.excl-reason` (línea 315) `border-radius:3px` → `99px`.

- [ ] **Step 4:** `.toast` (línea 203): `background:var(--bg3)` → `background:rgba(13,24,32,0.95)`, `box-shadow:0 8px 40px rgba(0,0,0,.7)` → `0 10px 30px rgba(0,0,0,0.5)`, agregar `backdrop-filter:blur(10px)`, `border-radius:var(--r)` → `10px`.

- [ ] **Step 5:** `.modal` (línea 187): `border-radius:var(--r)` → `14px`.

- [ ] **Step 6:** Commit: `git add index.html; git commit -m "style(admin): wizard con glow, badges pill, toast glass, topbar refinada"`

---

### Task 3: Avatares con iniciales + borrar renderTasa()

**Files:** Modify: `index.html` (JS)

- [ ] **Step 1:** En la sección UTILS (junto a `const esc=` línea ~3892) agregar:

```javascript
const inic=n=>String(n||"?").trim().split(/\s+/).map(w=>w[0]).filter(Boolean).slice(0,2).join("").toUpperCase();
```

- [ ] **Step 2:** Template de winner rows (línea 3162): después de `<div class="wr-n">${i+1}</div>` insertar `<div class="avi">${inic(g.nombre)}</div>`.

- [ ] **Step 3:** Templates `.mli` (líneas 3295 y 3317): envolver con avatar gris. Reemplazar el interior por:

```javascript
`<div class="mli" data-pid="${p.id}" onclick="selRemp(this.dataset.pid)" style="display:flex;align-items:center;gap:10px"><div class="avi gray">${inic(p.nombre)}</div><div style="flex:1;min-width:0"><div class="mli-n">${p.nombre}</div><div class="mli-e">${p.email}</div></div></div>`
```

(y el equivalente con `agregarGan` + `${badge}` en línea 3317).

- [ ] **Step 4:** Borrar la función `renderTasa()` completa (líneas 2568-2601). Verificado: nadie la llama y sus elementos `tasa-*` no existen en el HTML.

- [ ] **Step 5:** Verificación estática: extraer el `<script>` a temp y `node --check`. Expected: sin errores.

- [ ] **Step 6:** Commit: `git add index.html; git commit -m "style(admin): avatares con iniciales en listas; chore: borrar renderTasa muerta"`

---

### Task 4: Verificación visual (Mateo) y push

- [ ] `git push` (a pedido del usuario).
- [ ] Revisión visual en GitHub Pages: dashboard (cards con profundidad, 3 KPIs), wizard (paso activo con glow), listas con avatares, modal de reemplazo, toasts, historial.
