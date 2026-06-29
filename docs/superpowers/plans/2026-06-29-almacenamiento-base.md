# Etapa 1 — Base de almacenamiento confiable · Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que `localStorage` no se llene nunca (dejar de guardar históricos) y que un fallo de guardado se vea, manteniendo el anti-repetidos seguro.

**Architecture:** Cambios solo en `index.html` (app de una sola página, vanilla JS, sin build). Los históricos siguen viniendo de la planilla vía `importarDesdeTracking` (ya corre solo al abrir); se deja de duplicarlos en `localStorage`. Se agregan dos carteles fijos y una bandera de estado que bloquea el sorteo si los históricos no cargaron.

**Tech Stack:** HTML + JavaScript vanilla embebido en `index.html`. Sin framework de tests → **la verificación es por consola del navegador** (F12 → Console) en la app corriendo, que es el método de verificación ya usado en este proyecto.

## Global Constraints

- **Solo `index.html`.** No tocar `sorteo_script.gs` (Apps Script) → evita riesgo de CORS/deploy. [[project_apps_script_deploy]]
- **Rutas relativas** en cualquier HTML nuevo (deploy en GitHub Pages + Netlify). [[project_dos_deploys]]
- **No cambiar la fuente de verdad de históricos:** sigue siendo la hoja "Tracking Ganadores".
- **No introducir archivos JS externos** ni reestructurar el archivo (la app es un único `index.html` autocontenido).
- Spec de referencia: `docs/superpowers/specs/2026-06-29-almacenamiento-base-design.md`.

## Quién ejecuta qué

- **Agente:** edita `index.html` y hace los commits.
- **Mateo:** corre las verificaciones por consola en la app real (su navegador tiene los ~959 históricos y la sesión autenticada) y reporta el resultado. Cada tarea marca explícitamente los pasos "Verificá vos".

## Estructura de archivos

| Archivo | Responsabilidad | Cambio |
|---|---|---|
| `index.html` | App completa | `save()` (≈1265), markup de carteles tras `#toast` (1236), helpers de cartel, bandera `_historicosCargados` + gate de `ejecutarSorteo()` (≈3046) y `importarDesdeTracking()` (≈2454) |

> Las líneas son aproximadas y se corren al editar. Anclar por **nombre de función / id**, no por número de línea.

---

## Task 1: `save()` deja de persistir históricos

**Files:**
- Modify: `index.html` — función `save` (actualmente ≈1265-1283)

**Interfaces:**
- Consumes: `S.ganadores` (array; cada ganador puede tener `_historico: true`).
- Produces: `localStorage["ma_v6"].ganadores` contiene **solo** ganadores con `_historico` ausente/falsy. `S.ganadores` en memoria queda intacto (sigue con históricos).

- [ ] **Step 1 (Mateo — verificación previa, debe mostrar el problema):** En la app, F12 → Console:

```js
JSON.parse(localStorage.getItem("ma_v6")||"{}").ganadores.filter(g=>g._historico).length
```

Esperado AHORA: un número **mayor que 0** (cientos). Eso es el bug: los históricos están en el navegador.

- [ ] **Step 2 (Agente):** En `save()`, cambiar la línea de `ganadores` dentro de `toSave`.

Reemplazar:

```js
      ganadores:S.ganadores,
```

por:

```js
      ganadores:S.ganadores.filter(g=>!g._historico),
```

(No tocar el resto de `save()` en esta tarea; el `catch` se modifica en la Task 2.)

- [ ] **Step 3 (Agente):** Commit.

```bash
git add index.html
git commit -m "fix(persistencia): save() no guarda ganadores historicos en localStorage"
```

- [ ] **Step 4 (Mateo — verificación posterior, debe pasar):** Recargar la app, esperar ~3 s a que carguen los históricos, y en Console:

```js
save();
const g = JSON.parse(localStorage.getItem("ma_v6")).ganadores;
console.log("historicos en localStorage:", g.filter(x=>x._historico).length); // esperado: 0
console.log("historicos en memoria (anti-rep ok):", S.ganadores.some(x=>x._historico)); // esperado: true
console.log("tamano ma_v6 KB:", Math.round(localStorage.getItem("ma_v6").length/1024));
```

Esperado: `historicos en localStorage: 0`, `historicos en memoria: true`, y el tamaño cae a pocos KB.

---

## Task 2: Cartel visible cuando `save()` falla

**Files:**
- Modify: `index.html` — markup tras `#toast` (línea 1236); helpers nuevos en el `<script>`; `catch`/`try` de `save()`

**Interfaces:**
- Consumes: nada externo.
- Produces: `showSaveError()` y `clearSaveError()` (globales). `save()` llama a `clearSaveError()` al guardar OK y a `showSaveError()` en el `catch`.

- [ ] **Step 1 (Agente):** Agregar el cartel justo después de la línea `<div class="toast" id="toast"></div>` (1236):

```html
<div id="save-error-banner" style="display:none;position:fixed;top:0;left:0;right:0;z-index:10000;padding:10px 16px;background:#c0392b;color:#fff;font-size:13px;font-weight:600;text-align:center;box-shadow:0 2px 12px rgba(0,0,0,.4)">⚠ No se pudo guardar localmente. No cierres la pestaña hasta que se solucione.</div>
```

- [ ] **Step 2 (Agente):** Agregar los helpers en el `<script>` (junto a `toast`, ≈3885):

```js
function showSaveError(){const b=document.getElementById("save-error-banner");if(b)b.style.display="block";}
function clearSaveError(){const b=document.getElementById("save-error-banner");if(b)b.style.display="none";}
```

- [ ] **Step 3 (Agente):** Conectar `save()` a los helpers. Tras `localStorage.setItem("ma_v6",JSON.stringify(toSave));` agregar `clearSaveError();`, y en el `catch` agregar `showSaveError();`. El `catch` queda:

```js
  }catch(e){
    console.warn("localStorage save error:",e);
    showSaveError();
  }
```

y la línea de éxito:

```js
    localStorage.setItem("ma_v6",JSON.stringify(toSave));
    clearSaveError();
```

- [ ] **Step 4 (Agente):** Commit.

```bash
git add index.html
git commit -m "feat(persistencia): cartel persistente si falla el guardado local"
```

- [ ] **Step 5 (Mateo — verificación, rojo→verde):** En Console, forzar un fallo y verificar el cartel, después restaurar:

```js
const _orig = localStorage.setItem.bind(localStorage);
localStorage.setItem = () => { throw new DOMException("forzado","QuotaExceededError"); };
save();
console.log("cartel visible:", document.getElementById("save-error-banner").style.display); // esperado: "block"
localStorage.setItem = _orig;
save();
console.log("cartel oculto:", document.getElementById("save-error-banner").style.display); // esperado: "none"
```

Esperado: aparece el cartel rojo arriba al primer `save()`, y desaparece tras restaurar y volver a guardar.

---

## Task 3: Bloquear el sorteo si los históricos no cargaron

**Files:**
- Modify: `index.html` — markup tras `#toast`; bandera + funciones nuevas en `<script>`; `importarDesdeTracking()` (≈2454); `ejecutarSorteo()` (≈3046); `renderStep2()` (≈2976)

**Interfaces:**
- Consumes: `_historicosCargados` (boolean global, init `false`); `S.eventos`, `S.evActivo`, `ev.antiRep`.
- Produces: `refreshSortGate()` (habilita/deshabilita `#btn-sort` y muestra/oculta `#hist-block-banner`); `reintentarHistoricos()`. `importarDesdeTracking` setea `_historicosCargados` y llama a `refreshSortGate()`.

- [ ] **Step 1 (Agente):** Agregar el cartel de bloqueo después del `#save-error-banner` (Task 2):

```html
<div id="hist-block-banner" style="display:none;position:fixed;top:0;left:0;right:0;z-index:9999;padding:10px 16px;background:#b9770e;color:#fff;font-size:13px;font-weight:600;text-align:center">⚠ No se pudieron cargar los ganadores históricos — el anti-repetidos no funciona. <button onclick="reintentarHistoricos()" style="margin-left:10px;padding:4px 12px;border:1px solid #fff;background:transparent;color:#fff;border-radius:6px;cursor:pointer;font-weight:700">Reintentar</button></div>
```

- [ ] **Step 2 (Agente):** Agregar la bandera y funciones en el `<script>` (cerca de la STATE / `save`):

```js
let _historicosCargados=false;
function reintentarHistoricos(){importarDesdeTracking(false);}
function refreshSortGate(){
  const ev=S.eventos.find(e=>e.id===S.evActivo);
  const needHist=ev&&ev.antiRep!==false;
  const blocked=needHist&&!_historicosCargados;
  const btn=document.getElementById("btn-sort");
  const banner=document.getElementById("hist-block-banner");
  if(btn)btn.disabled=blocked;
  if(banner)banner.style.display=blocked?"block":"none";
}
```

- [ ] **Step 3 (Agente):** En `importarDesdeTracking(silente)`, setear la bandera. Al **inicio** del `try` (antes de `const res=await api(...)`):

```js
    _historicosCargados=false;
```

En las salidas tempranas por error (`if(!res.ok){...return;}` y `if(!cols.length){...return;}`), antes del `return` agregar:

```js
      refreshSortGate();
```

Tras reconstruir todo OK (después de `if(res.ticketsBase>0)...` / al final del bloque de éxito, antes del `if(!silente)toast(...)` de la línea ≈2541):

```js
    _historicosCargados=true;
    refreshSortGate();
```

En el `catch` de la función, antes de su cierre, agregar `refreshSortGate();`.

- [ ] **Step 4 (Agente):** En `ejecutarSorteo()` (≈3046), agregar el guard justo después de `if(!ev)return;`:

```js
  if(ev.antiRep!==false && !_historicosCargados){
    refreshSortGate();
    toast("No se pudieron cargar los históricos — el anti-repetidos no funciona. Reintentá.","warn");
    return;
  }
```

- [ ] **Step 5 (Agente):** Asegurar que el gate se refresca al entrar al paso 2. Al final de `renderStep2()` (≈2976+) agregar:

```js
  refreshSortGate();
```

- [ ] **Step 6 (Agente):** Commit.

```bash
git add index.html
git commit -m "feat(sorteo): bloquear sorteo si no cargaron los historicos (anti-rep seguro)"
```

- [ ] **Step 7 (Mateo — verificación):** En Console, simular históricos no cargados con un evento que use anti-repetición (asegurate de que el evento activo tenga anti-rep activado):

```js
_historicosCargados=false; refreshSortGate();
console.log("btn deshabilitado:", document.getElementById("btn-sort").disabled); // esperado: true
console.log("banner visible:", document.getElementById("hist-block-banner").style.display); // esperado: "block"
ejecutarSorteo();
console.log("sorteo bloqueado (no arrancó el bombo)"); // verificá visualmente que NO empezó a girar
```

Después, reintentar la carga real:

```js
reintentarHistoricos();
// esperar ~2 s
setTimeout(()=>{ console.log("cargados:",_historicosCargados,"| btn:",document.getElementById("btn-sort").disabled); }, 2500);
```

Esperado: con históricos cargados → `cargados: true`, `btn: false`, banner oculto.

- [ ] **Step 8 (Mateo — verificación caso sin anti-rep):** Con un evento que tenga `antiRep === false` activo:

```js
_historicosCargados=false; refreshSortGate();
console.log("btn deshabilitado:", document.getElementById("btn-sort").disabled); // esperado: false (no bloquea)
```

---

## Verificación de cierre (Etapa 1 completa, Mateo)

Tras las 3 tareas, hacer un sorteo de prueba completo y un envío chico, después:

```js
const raw = localStorage.getItem("ma_v6");
const g = JSON.parse(raw).ganadores;
console.log("KB:", Math.round(raw.length/1024), "| historicos persistidos:", g.filter(x=>x._historico).length);
```

Esperado: tamaño chico (pocas decenas de KB) y `historicos persistidos: 0`. El historial, el conteo y las exclusiones "Ya ganó 🏆" se ven igual que antes.

## Self-review (hecho por el agente al escribir el plan)

- **Cobertura del spec:** Cambio 1 → Task 1; Cambio 2 → Task 2; Cambio 3 → Task 3; caso `antiRep===false` → Task 3 Step 4/8; criterios de éxito → Verificación de cierre. ✓
- **Sin placeholders:** todos los pasos tienen código real. ✓
- **Consistencia de nombres:** `_historicosCargados`, `refreshSortGate`, `reintentarHistoricos`, `showSaveError`, `clearSaveError`, ids `save-error-banner` / `hist-block-banner` usados de forma idéntica en todas las tareas. ✓
