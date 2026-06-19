# Envío robusto multiusuario (defectos 3 + 2) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que el botón "Enviar via Gmail" distinga claramente "nada que enviar" de "desincronización" (defecto 3) y que el envío use la misma clave canónica con la que se escriben los ganadores, para que funcione desde cualquier computadora (defecto 2).

**Architecture:** Todo el cambio vive en el frontend (`index.html`). Defecto 3: lógica de mensajes según los conteos locales del evento. Defecto 2: el envío y la búsqueda de PDFs pasan a usar `ev.id` (ID canónico de la hoja Shows, igual en todas las compus) en vez de `ev.nombre`. Sin migración de datos; los ganadores ya se escriben con `ev.id`.

**Tech Stack:** HTML/JS vanilla (un solo `index.html`), backend Google Apps Script (sin cambios). Sin framework de tests — verificación = `node --check` del bloque `<script>` extraído + prueba manual en navegador con stub de `window.api`.

**Spec:** `docs/superpowers/specs/2026-06-19-envio-robusto-multiusuario-design.md`

## Global Constraints

- Un solo archivo de UI: `index.html`. No hay build ni bundler.
- No tocar el backend (`sorteo_script.gs`).
- No cambiar el flujo de envío cuando sí hay pendientes y los identificadores coinciden.
- `toast(msg,type)` solo tiene dos estilos: `type="ok"` (✓ verde, default) y cualquier otro valor (⚠). Definido en `index.html:3868`.
- Mantener el estilo del archivo (arrow functions, comillas dobles, sin reformatear código aledaño). No agregar dependencias.
- Orden obligatorio: Task 1 y 2 (defecto 3) antes de Task 3 (defecto 2). Cada task asume el estado dejado por la anterior.

---

### Task 1: (defecto 3) Distinguir "sin ganadores" de "todos enviados" en `enviarGmail`

**Files:**
- Modify: `index.html` — función `enviarGmail`, branch `if(!pend.length)` (`index.html:3777-3778`)

**Interfaces:**
- Consumes: `S.ganadores`, `S.evActivo` (`evId`), `toast(msg,type)`.
- Produces: ningún símbolo nuevo; solo cambia el mensaje del branch sin pendientes.

- [ ] **Step 1: Reproducir el comportamiento actual (verificación previa)**

Abrir `index.html` en el navegador. Ir a un evento sin ganadores y otro con todos enviados, apretar "Enviar via Gmail" en cada uno: en ambos aparece el mismo gris "No hay pendientes". Ese es el comportamiento a mejorar.

- [ ] **Step 2: Reemplazar el branch sin pendientes**

Localizar (`index.html:3777-3778`):

```javascript
  const pend=S.ganadores.filter(g=>g.evId===evId&&g.estado==="pendiente");
  if(!pend.length){toast("No hay pendientes","warn");return;}
```

Reemplazar por:

```javascript
  const pend=S.ganadores.filter(g=>g.evId===evId&&g.estado==="pendiente");
  if(!pend.length){
    const totalEv=S.ganadores.filter(g=>g.evId===evId).length;
    if(!totalEv){
      toast("Este show todavía no tiene ganadores sorteados.","warn");
    } else {
      toast("Ya enviaste las entradas a los "+totalEv+" ganadores de este show.");
    }
    return;
  }
```

- [ ] **Step 3: Verificar sintaxis**

Extraer el contenido del bloque `<script>` principal de `index.html` a un archivo temporal `.js` y correr:

Run: `node --check <archivo-temporal>.js`
Expected: sin errores de sintaxis.

- [ ] **Step 4: Verificar comportamiento (navegador)**

Evento sin ganadores → "Este show todavía no tiene ganadores sorteados." (⚠). Evento con todos los ganadores en `estado:"enviado"` → "Ya enviaste las entradas a los N ganadores de este show." (✓ verde, N correcto).

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat(envio): distinguir 'sin ganadores' de 'todos enviados' al no haber pendientes"
```

---

### Task 2: (defecto 3) Detectar desincronización en `_enviarMailsCore`

**Files:**
- Modify: `index.html` — función `_enviarMailsCore` (`index.html:3798-3808`)

**Interfaces:**
- Consumes: respuesta del backend `res` (`{ok, enviados, enviadosMails, errores, mensaje}`), `pend` (array de ganadores que se intentó enviar).
- Produces: ningún símbolo nuevo.

- [ ] **Step 1: Reproducir el caso (verificación previa)**

En la consola del navegador, stubear el backend para simular "0 pendientes en la planilla" con pendientes locales:

```js
const _origApi = window.api;
window.api = async (p) => { if (p.action === "enviarMails") return {ok:true, enviados:0, enviadosMails:[], errores:[], mensaje:"No hay ganadores pendientes"}; if (p.action === "checkPDFs") return {ok:true, status:"ok", found:99, expected:1}; return _origApi(p); };
```

Con un evento que tenga ganadores pendientes locales, apretar "Enviar". Hoy muestra un toast tipo éxito con "No hay ganadores pendientes" — indistinguible de un envío real. Restaurar con `window.api=_origApi` al terminar.

- [ ] **Step 2: Agregar la detección del caso (c)**

Localizar en `_enviarMailsCore` (`index.html:3803`):

```javascript
  if(!res.ok){toast("Error: "+(res.error||"sin respuesta"),"warn");return;}
```

Insertar INMEDIATAMENTE DESPUÉS de esa línea:

```javascript
  // Desincronización: esperábamos mandar (pend>0) pero la planilla no devolvió ningún
  // pendiente y no hubo errores de envío → las filas no están en Sheets para este show.
  if(res.enviados===0 && !(res.errores&&res.errores.length) && pend.length>0){
    alert("Esperaba "+pend.length+" ganador(es) pendiente(s) para este show, pero la planilla no devolvió ninguno.\n\nPosible desincronización — revisá/reintentá la sincronización antes de enviar.");
    return;
  }
```

- [ ] **Step 3: Verificar sintaxis**

Run: `node --check <archivo-temporal>.js`
Expected: sin errores.

- [ ] **Step 4: Verificar comportamiento (navegador)**

Con el stub del Step 1 y un evento con pendientes locales, apretar "Enviar" → aparece el `alert` de desincronización y NO se marca ningún ganador como enviado. Quitar el stub y hacer un envío normal (o un stub con `enviados:>0`) → flujo normal sin el alert.

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat(envio): alertar desincronización cuando el backend devuelve 0 pendientes con pendientes locales"
```

---

### Task 3: (defecto 2) Usar `ev.id` como clave de envío y búsqueda de PDFs

**Files:**
- Modify: `index.html` — función `enviarGmail` (`index.html:3779-3796`) y función `_enviarMailsCore` (firma + llamada a `enviarMails`, `index.html:3798-3801`)

**Interfaces:**
- Consumes: `ev` (objeto evento con `.id`, `.nombre`, `.entradasXGan`), `evId` (= `S.evActivo`), `pend`, `api`, `mostrarPreflightPDFs(check,pend,ev,nombreParaDrive)`.
- Produces: `_enviarMailsCore(showKey, showName, ev, pend)` — `showKey` = `ev.id` (clave para matchear filas de Ganadores en el backend); `showName` = `ev.nombre` (nombre para la carpeta de Drive y la UI).

**Contexto (por qué):** la escritura de ganadores usa `ev.id` como `showId` (`index.html:3582`), pero el envío usa `ev.nombre`. Cuando difieren (46 de 71 shows según la auditoría, y siempre tras recargar desde la nube en otra compu), el backend no matchea y devuelve 0 enviados. Alinear el envío a `ev.id` lo arregla sin tocar datos. La búsqueda de PDFs sigue funcionando porque el backend matchea la carpeta por `showNombre`, luego `showId`, luego substring (`sorteo_script.gs:566-588`), y le seguimos pasando `ev.nombre` como `showNombre`.

- [ ] **Step 1: Reproducir el bug (verificación previa)**

En consola, simular un evento cargado desde la nube con `id ≠ nombre` (el caso multi-compu) y observar qué `showId` se manda hoy:

```js
const _origApi = window.api;
window.api = async (p) => { if (p.action === "checkPDFs"){ console.log("checkPDFs showId=", p.showId); return {ok:true,status:"ok",found:99,expected:1}; } if (p.action === "enviarMails"){ console.log("enviarMails showId=", p.showId); return {ok:true,enviados:0,enviadosMails:[],errores:[],mensaje:"No hay ganadores pendientes"}; } return _origApi(p); };
```

Tomar un evento cuyo `ev.id` ≠ `ev.nombre` (por consola: `S.eventos.find(e=>e.id!==e.nombre)` y setear `S.evActivo` a su `id`, agregarle un ganador pendiente local), apretar "Enviar". En consola se ve que hoy `enviarMails showId=` es el **nombre**, no el id. Restaurar `window.api=_origApi`.

- [ ] **Step 2: Modificar `enviarGmail` para separar clave (id) y nombre (Drive/UI)**

Localizar en `enviarGmail` (`index.html:3779-3796`):

```javascript
  const nombreCarpeta=ev?.nombre||evId;
  const btn=document.getElementById("btn-gmail");
  btn.disabled=true;btn.textContent="🔎 Verificando PDFs en Drive…";
  const check=await api({
    action:"checkPDFs",
    showId:String(nombreCarpeta),
    showNombre:nombreCarpeta,
    entradasXGan:ev?.entradasXGan||1,
    pendCount:pend.length
  });
  btn.disabled=false;btn.textContent="Enviar via Gmail";
  if(!check.ok){toast("Error verificando PDFs: "+(check.error||""),"warn");return;}
  if(check.status!=="ok"){
    mostrarPreflightPDFs(check,pend,ev,nombreCarpeta);
    return;
  }
  if(!confirm("✓ Verificado: "+check.found+" PDF(s) en Drive (necesarios: "+check.expected+").\n\n¿Enviar mails a "+pend.length+" ganador(es) de \""+nombreCarpeta+"\"?"))return;
  await _enviarMailsCore(nombreCarpeta,ev,pend);
```

Reemplazar por:

```javascript
  // showKey: clave canónica con la que se ESCRIBEN los ganadores (ev.id) — la usa el backend
  // para matchear filas. showName: nombre para la carpeta de Drive y los textos de la UI.
  const showKey=ev?.id||evId;
  const showName=ev?.nombre||evId;
  const btn=document.getElementById("btn-gmail");
  btn.disabled=true;btn.textContent="🔎 Verificando PDFs en Drive…";
  const check=await api({
    action:"checkPDFs",
    showId:String(showKey),
    showNombre:showName,
    entradasXGan:ev?.entradasXGan||1,
    pendCount:pend.length
  });
  btn.disabled=false;btn.textContent="Enviar via Gmail";
  if(!check.ok){toast("Error verificando PDFs: "+(check.error||""),"warn");return;}
  if(check.status!=="ok"){
    mostrarPreflightPDFs(check,pend,ev,showName);
    return;
  }
  if(!confirm("✓ Verificado: "+check.found+" PDF(s) en Drive (necesarios: "+check.expected+").\n\n¿Enviar mails a "+pend.length+" ganador(es) de \""+showName+"\"?"))return;
  await _enviarMailsCore(showKey,showName,ev,pend);
```

- [ ] **Step 3: Modificar la firma de `_enviarMailsCore` y la llamada a `enviarMails`**

Localizar (`index.html:3798-3801`, ya con la detección de desincronización de Task 2 presente debajo):

```javascript
async function _enviarMailsCore(nombreCarpeta,ev,pend){
  const btn=document.getElementById("btn-gmail");
  btn.disabled=true;btn.textContent="⏳ Enviando…";
  const res=await api({action:"enviarMails",showId:String(nombreCarpeta),entradasXGan:ev?.entradasXGan||1});
```

Reemplazar por:

```javascript
async function _enviarMailsCore(showKey,showName,ev,pend){
  const btn=document.getElementById("btn-gmail");
  btn.disabled=true;btn.textContent="⏳ Enviando…";
  const res=await api({action:"enviarMails",showId:String(showKey),entradasXGan:ev?.entradasXGan||1});
```

(El parámetro `showName` queda disponible para textos futuros; el cuerpo restante de `_enviarMailsCore` —incluida la detección de desincronización de Task 2— no cambia, ya que no usaba `nombreCarpeta` salvo en la llamada a `enviarMails`.)

- [ ] **Step 4: Verificar sintaxis**

Run: `node --check <archivo-temporal>.js`
Expected: sin errores.

- [ ] **Step 5: Verificar comportamiento (navegador)**

Repetir el stub del Step 1. Ahora la consola debe mostrar `checkPDFs showId=` y `enviarMails showId=` con el **`ev.id`** (no el nombre). Confirmar que para un evento normal donde `id===nombre` el comportamiento es idéntico al anterior.

- [ ] **Step 6: Verificar que no hay evento "fantasma" que pise el id canónico (lectura)**

Leer el bloque de reconstrucción de `importarTracking` (`index.html:2481-2502`). Confirmar que: antes de crear un evento `_fromTracking`, busca un evento existente que matchee por `show`/`nombre` + fecha (`index.html:2483-2487`), de modo que un show ya cargado desde la nube NO se duplica con otro id. Documentar en el reporte: el envío usa `S.evActivo` (evento de la lista de shows, proveniente del cloud vía `syncShowsFromCloud`, `index.html:3911`), y el backend matchea las filas de Ganadores por el `showId` enviado (`ev.id`), independientemente de los objetos ganador locales — así que un eventual evento `_fromTracking` solo afecta la vista de historial, no el envío. Si la lectura revela un camino donde `S.evActivo` puede quedar apuntando a un evento `_fromTracking` con id distinto al del cloud para un show activo, reportarlo como hallazgo (no arreglar en esta task).

- [ ] **Step 7: Commit**

```bash
git add index.html
git commit -m "fix(envio): usar ev.id (clave canónica) en envío y checkPDFs para que funcione multi-compu"
```

---

### Task 4: Verificación integral + limpieza de la entrada duplicada

**Files:**
- Sin cambios de código. Verificación manual end-to-end + limpieza de datos en la hoja Shows.

- [ ] **Step 1: Recorrido completo de los tres casos del defecto 3**

(a) evento con todos enviados → toast verde "ya enviaste a los N"; (b) evento sin sortear → toast ⚠ "no tiene ganadores sorteados"; (c) con el stub de `enviarMails`→0 y pendientes locales → `alert` de desincronización sin marcar enviados.

- [ ] **Step 2: Escenario multi-compu del defecto 2**

Tomar un show con `ev.id ≠ ev.nombre`. Verificar (consola/Network) que `enviarMails` y `checkPDFs` ahora mandan `ev.id`. Si hay un entorno de prueba seguro, confirmar contra la planilla real que las filas de Ganadores guardadas con ese `ev.id` ahora matchean y se envían.

- [ ] **Step 3: Limpiar la entrada duplicada de Angela Leiva**

En la hoja **Shows**: confirmar con Mateo cuál entrada de Angela Leiva conservar (`ANGELA LEIVA` activa cupo 15 vs `2026-06-18 ANGELA LEIVA` ya Eliminado). Dejar solo la correcta. Operación manual acotada — no migración masiva.

- [ ] **Step 4: Actualizar la memoria del proyecto**

Marcar en `project-bug-envio-angela-leiva` que los defectos 3 y 2 quedaron resueltos por este plan (enfoque A, sin migración), y que la entrada duplicada fue limpiada.

---

## Notas

- Este plan no toca el backend. El defecto 3 (opcional, mínimo) podría mejorar el `mensaje` de `sorteo_script.gs:340`, pero se omite porque el frontend ya distingue los tres casos.
- Sin regresión esperada en el defecto 2: para shows donde `id===nombre` el envío es idéntico; para shows donde difieren, hoy el envío está roto y este cambio lo arregla.
