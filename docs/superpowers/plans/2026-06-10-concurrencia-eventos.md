# Concurrencia entre usuarios al editar eventos — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que dos personas de RRHH editando eventos en paralelo no se pisen cambios en silencio. Optimistic locking server-side + poller silencioso cliente-side + banner sutil en el editor cuando algo remoto afecta lo que está abierto.

**Architecture:** Versionado por timestamp (`ActualizadoEn` de la pestaña `Shows`) escrito por el servidor. `upsertShow` opcionalmente verifica el `expectedVersion` que mandó el cliente y rechaza con `conflict` si no coincide. Cliente guarda dos timestamps por evento: `_version` (última versión cloud conocida, mutable) y `_versionAtOpen` (snapshot al abrir el editor, inmutable hasta cerrar). Poller cada 60s gateado por `document.visibilityState`. Banner contextual sobre el editor cuando el evento abierto cambió o fue borrado.

**Tech Stack:** Google Apps Script (sin TypeScript), HTML/JS vanilla en un solo archivo, sin build step, sin tests automatizados (verificación manual definida en el spec).

**Spec de referencia:** [docs/superpowers/specs/2026-06-10-concurrencia-eventos-design.md](../specs/2026-06-10-concurrencia-eventos-design.md)

**Contexto operativo:**
- Después de editar `sorteo_script.gs` el cambio NO está vivo hasta hacer un deploy nuevo del Apps Script. El deploy modo "Ejecutar como: Yo + Cualquier persona" no debe cambiarse (memoria de proyecto — rompe CORS).
- La app del admin se sirve desde GitHub Pages (push a `main` la deploya automáticamente). Para iterar rápido, conviene abrir `index.html` localmente o probar contra una copia del Apps Script en deploy de test antes de publicar.
- No hay tests automatizados. Cada tarea tiene una verificación manual concreta.

---

## File Structure

**Modificar:**
- `sorteo_script.gs` — funciones `getShowsCloud` (~líneas 1047-1090) y `upsertShow` (~líneas 1092-1130).
- `index.html`:
  - JS: `syncShowsFromCloud` (~3798), `abrirEdit` (~1409), `guardarEvento` (~1426), `eliminarEvento` (~1455), `crearEvento` handler (~1387), `startApp` (~3825). Función nueva `pollShowsLight`. Handlers nuevos para banner y modal de conflicto.
  - CSS: estilos para banner y modal de conflicto, dentro del `<style>` que ya existe.
  - HTML: bloque del banner dentro de `modal-edit`, bloque del modal de conflicto al final del `<body>`.

**No tocar:**
- `formulario_inscripcion.html`, lógica de inscripciones, sorteo, ganadores, mails, tracking.
- Función `deleteShow` del Apps Script.
- El dispatcher `doPost` (el `expectedVersion` viaja como propiedad del objeto `show`, no como param suelto).

---

## Task 1: Backend — `getShowsCloud` devuelve `actualizadoEn`

**Files:**
- Modify: `sorteo_script.gs:1066-1083` (loop dentro de `getShowsCloud`)

- [ ] **Step 1: Agregar `actualizadoEn` al objeto del show devuelto**

En `sorteo_script.gs`, dentro del `for` de `getShowsCloud` (líneas ~1066-1083), agregar una propiedad al objeto pusheado:

```javascript
for (let i = 1; i < datos.length; i++) {
  if (String(datos[i][12]).trim() === "1") continue; // eliminado
  const id = String(datos[i][0]).trim();
  if (!id) continue;
  shows.push({
    id:            id,
    show:          String(datos[i][1] || "").trim(),
    nombre:        String(datos[i][2] || "").trim(),
    fecha:         formatFecha(datos[i][3]),
    hora:          String(datos[i][4] || "").trim(),
    venue:         String(datos[i][5] || "").trim(),
    cantidad:      parseInt(datos[i][6]) || 2,
    entradasXGan:  parseInt(datos[i][7]) || 1,
    formUrl:       String(datos[i][8] || "").trim(),
    antiRep:       String(datos[i][9]).trim() !== "0",
    creadoEn:      String(datos[i][10] || "").trim(),
    actualizadoEn: String(datos[i][11] || "").trim(),
  });
}
```

Nota: índice 11 = columna 12 (`ActualizadoEn`). Está garantizado por el header en `_ensureShowsSheet`.

- [ ] **Step 2: Probar la función desde el editor de Apps Script**

En el editor de Apps Script, seleccionar la función `getShowsCloud` y ejecutarla. Abrir Logger (Ver → Registros).

Expected: el log dice `getShowsCloud: N shows devueltos.` y no hay error.

Para verificar el campo nuevo, agregar temporalmente al final de la función (solo para esta prueba, después se borra):

```javascript
Logger.log(JSON.stringify(shows[0], null, 2));
```

Expected: el objeto loggeado incluye `"actualizadoEn": "DD/MM/YYYY HH:mm:ss"`.

Borrar el `Logger.log` temporal.

- [ ] **Step 3: Commit**

```bash
git add sorteo_script.gs
git commit -m "feat(backend): getShowsCloud devuelve actualizadoEn por show"
```

---

## Task 2: Backend — `upsertShow` valida `expectedVersion`

**Files:**
- Modify: `sorteo_script.gs:1092-1130` (función completa `upsertShow`)

- [ ] **Step 1: Reescribir `upsertShow` con chequeo de versión opcional**

Reemplazar la función `upsertShow` completa por:

```javascript
function upsertShow(show) {
  try {
    if (!show || !show.id) return { ok: false, error: "show.id requerido" };
    const hoja = _ensureShowsSheet();
    const datos = hoja.getDataRange().getValues();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    const expectedVersion = show.expectedVersion || null;
    const row = [
      String(show.id),
      String(show.show || ""),
      String(show.nombre || ""),
      String(show.fecha || ""),
      String(show.hora || ""),
      String(show.venue || ""),
      parseInt(show.cantidad) || 2,
      parseInt(show.entradasXGan) || 1,
      String(show.formUrl || ""),
      show.antiRep !== false ? "1" : "0",
      String(show.creadoEn || now),
      now,
      "0"
    ];
    // Buscar si ya existe
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === String(show.id).trim()) {
        const currentVersion = String(datos[i][11] || "").trim();
        // Chequeo de optimistic locking
        if (expectedVersion && expectedVersion !== currentVersion) {
          // Construir el show actual del cloud para devolverlo
          const currentShow = {
            id:            String(datos[i][0]).trim(),
            show:          String(datos[i][1] || "").trim(),
            nombre:        String(datos[i][2] || "").trim(),
            fecha:         String(datos[i][3] || "").trim(),
            hora:          String(datos[i][4] || "").trim(),
            venue:         String(datos[i][5] || "").trim(),
            cantidad:      parseInt(datos[i][6]) || 2,
            entradasXGan:  parseInt(datos[i][7]) || 1,
            formUrl:       String(datos[i][8] || "").trim(),
            antiRep:       String(datos[i][9]).trim() !== "0",
            creadoEn:      String(datos[i][10] || "").trim(),
            actualizadoEn: currentVersion,
          };
          Logger.log("upsertShow (conflict): " + show.id + " expected=" + expectedVersion + " current=" + currentVersion);
          return { ok: false, error: "conflict", currentVersion: currentVersion, currentShow: currentShow };
        }
        // Conservar creadoEn original
        row[10] = String(datos[i][10] || now);
        hoja.getRange(i + 1, 1, 1, 13).setValues([row]);
        Logger.log("upsertShow (actualizado): " + show.id);
        return { ok: true, accion: "actualizado", newVersion: now };
      }
    }
    hoja.appendRow(row);
    Logger.log("upsertShow (creado): " + show.id);
    return { ok: true, accion: "creado", newVersion: now };
  } catch(e) {
    Logger.log("Error en upsertShow: " + e.message);
    return { ok: false, error: e.message };
  }
}
```

Cambios clave respecto a la versión actual:
1. Lee `expectedVersion` de `show.expectedVersion` (puede venir `undefined`/`null`).
2. Cuando hay match de `id` y `expectedVersion` está presente, compara contra `datos[i][11]` (columna `ActualizadoEn` actual).
3. Si no coincide → devuelve `{ok:false, error:"conflict", currentVersion, currentShow}` SIN modificar nada.
4. En éxito devuelve `newVersion: now` (el timestamp que se escribió).
5. `formatFecha` no se necesita acá — devolvemos la fecha como string raw del Sheet en el `currentShow`. El frontend ya tiene la fecha del cloud reciente del `getShowsCloud` que ejecuta el poller; este `currentShow` es solo para mostrar en el modal de conflicto.

- [ ] **Step 2: Probar conflict path con consola de Apps Script**

Agregar una función temporal al final del `.gs` para test:

```javascript
function _testConflict() {
  // Asume que existe al menos un show. Tomar el primero.
  const cloud = getShowsCloud();
  if (!cloud.ok || !cloud.shows.length) {
    Logger.log("Sin shows para testear. Creá uno desde la app primero.");
    return;
  }
  const target = cloud.shows[0];
  Logger.log("Show de prueba: " + target.id + " versión actual: " + target.actualizadoEn);

  // Caso 1: sin expectedVersion → debería actualizar OK
  const r1 = upsertShow({ ...target, cantidad: target.cantidad });
  Logger.log("Sin expectedVersion: " + JSON.stringify(r1));

  // Caso 2: con expectedVersion correcta → debería actualizar OK
  const cloud2 = getShowsCloud();
  const target2 = cloud2.shows.find(s => s.id === target.id);
  const r2 = upsertShow({ ...target2, expectedVersion: target2.actualizadoEn });
  Logger.log("Con expectedVersion correcta: " + JSON.stringify(r2));

  // Caso 3: con expectedVersion vieja → debería dar conflict
  const r3 = upsertShow({ ...target2, expectedVersion: "01/01/2020 00:00:00" });
  Logger.log("Con expectedVersion vieja: " + JSON.stringify(r3));
}
```

Ejecutar `_testConflict` desde el editor de Apps Script.

Expected (en el Logger):
- Caso 1: `{"ok":true,"accion":"actualizado","newVersion":"..."}`
- Caso 2: `{"ok":true,"accion":"actualizado","newVersion":"..."}`
- Caso 3: `{"ok":false,"error":"conflict","currentVersion":"...","currentShow":{...}}`

- [ ] **Step 3: Borrar la función `_testConflict`**

Eliminar la función `_testConflict` del archivo. Era solo para verificar.

- [ ] **Step 4: Commit**

```bash
git add sorteo_script.gs
git commit -m "feat(backend): upsertShow con optimistic locking por ActualizadoEn"
```

---

## Task 3: Deploy del Apps Script

**Files:** ninguno (acción manual)

- [ ] **Step 1: Hacer nuevo deploy del Apps Script**

En el editor de Apps Script: `Deploy` → `Manage deployments` → clic en el lápiz del deploy activo → cambiar `Version` a `New version` → descripción: "Optimistic locking eventos" → `Deploy`.

CRÍTICO: dejar `Execute as: Me` y `Who has access: Anyone` exactamente como estaban (memoria de proyecto: cambiar esto rompe CORS con el frontend).

- [ ] **Step 2: Verificar que la URL del deploy NO cambió**

Comparar la URL del deploy con la que está en la constante `SCRIPT_URL` del `index.html`. Tienen que coincidir. Si no coinciden, el deploy se hizo como "nuevo deploy" en vez de "nueva versión del deploy existente" — rehacer.

- [ ] **Step 3: Smoke test desde una pestaña incógnito**

Abrir un browser en incógnito y entrar a la app del admin. Abrir DevTools → Network. Recargar. Buscar la request a Apps Script y verificar que:
- La response 200 OK.
- El payload de `getShowsCloud` incluye `actualizadoEn` en cada show.

Si la response 401/redirige a login → el modo del deploy cambió por error; rehacer paso 1 con `Anyone`.

---

## Task 4: Frontend — poblar `_version` desde el cloud

**Files:**
- Modify: `index.html:3798-3824` (función `syncShowsFromCloud`)

- [ ] **Step 1: Mapear `actualizadoEn` a `_version` en `syncShowsFromCloud`**

En `index.html`, dentro de `syncShowsFromCloud`, después de la línea `S.eventos=res.shows;`, agregar el mapeo:

```javascript
async function syncShowsFromCloud(){
  try{
    const res=await api({action:"getShowsCloud"});
    if(!res.ok||!Array.isArray(res.shows))return;
    const cloudIds=new Set(res.shows.map(s=>s.id));
    // Subir shows locales que aún no están en la nube
    for(const ev of S.eventos){
      if(!cloudIds.has(ev.id)){
        api({action:"upsertShow",show:ev}).catch(()=>{});
      }
    }
    if(!res.shows.length)return; // No sobreescribir locales si la nube está vacía
    // Merge: la nube tiene prioridad, pero conservar ganadores/sorteo locales
    S.eventos=res.shows.map(s=>({...s,_version:s.actualizadoEn||""}));
    S.eventos.sort((a,b)=>b.fecha.localeCompare(a.fecha));
    // Validar que evActivo aún existe
    if(S.evActivo&&!S.eventos.find(e=>e.id===S.evActivo)){
      S.evActivo=S.eventos.length?S.eventos[0].id:null;
    }
    save();
    renderDash();
    updateTopbar();
    console.log("syncShowsFromCloud: "+res.shows.length+" shows sincronizados.");
  }catch(e){
    console.warn("syncShowsFromCloud error:",e);
  }
}
```

La única línea cambiada es `S.eventos=res.shows.map(s=>({...s,_version:s.actualizadoEn||""}));` (antes era `S.eventos=res.shows;`).

- [ ] **Step 2: Probar en browser**

Recargar el admin. Abrir DevTools → Console. Esperar ~1 segundo (sync inicial). Tipear:

```javascript
S.eventos.slice(0,3).map(e=>({id:e.id,_version:e._version}))
```

Expected: array de objetos donde cada `_version` es un string `"DD/MM/YYYY HH:mm:ss"` (no string vacío, no `undefined`).

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: poblar S.eventos[]._version desde actualizadoEn del cloud"
```

---

## Task 5: Frontend — snapshot `_versionAtOpen` al abrir el editor

**Files:**
- Modify: `index.html:1409-1425` (función `abrirEdit`)

- [ ] **Step 1: Setear `_versionAtOpen` en `abrirEdit`**

Reemplazar `abrirEdit` por:

```javascript
function abrirEdit(id){
  const ev=S.eventos.find(e=>e.id===id);
  if(!ev)return;
  // Snapshot de la versión cloud al abrir — usado como expectedVersion al guardar
  ev._versionAtOpen=ev._version||"";
  document.getElementById("edit-id").value=id;
  document.getElementById("edit-show").value=ev.show||"";
  document.getElementById("edit-nombre").value=ev.nombre||"";
  document.getElementById("edit-fecha").value=ev.fecha||"";
  document.getElementById("edit-hora").value=ev.hora||"";
  document.getElementById("edit-cant").value=ev.cantidad||2;
  document.getElementById("edit-exg").value=ev.entradasXGan||1;
  document.getElementById("edit-venue").value=ev.venue||"";
  document.getElementById("edit-form").value=ev.formUrl||"";
  document.getElementById("edit-anti").checked=ev.antiRep!==false;
  const slug=ev.id.toUpperCase().replace(/[^A-Z0-9]/g,"-").replace(/-+/g,"-").replace(/^-|-$/g,"");
  document.getElementById("edit-closeat").value=S.schedules[slug]||"";
  document.getElementById("modal-edit").classList.add("show");
}
```

Única línea agregada: `ev._versionAtOpen=ev._version||"";` (justo después del `if(!ev)return;`).

- [ ] **Step 2: Probar en browser**

Recargar admin. Abrir el editor de cualquier evento. En consola:

```javascript
const id=document.getElementById("edit-id").value;
const ev=S.eventos.find(e=>e.id===id);
console.log({_version:ev._version,_versionAtOpen:ev._versionAtOpen});
```

Expected: ambos campos tienen el mismo string `"DD/MM/YYYY HH:mm:ss"`.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: snapshot _versionAtOpen al abrir editor de evento"
```

---

## Task 6: Frontend — `guardarEvento` espera respuesta y manda `expectedVersion`

**Files:**
- Modify: `index.html:1426-1454` (función `guardarEvento`)

- [ ] **Step 1: Reescribir `guardarEvento` como async con manejo de conflict**

Reemplazar `guardarEvento` por:

```javascript
async function guardarEvento(){
  const id=document.getElementById("edit-id").value;
  const ev=S.eventos.find(e=>e.id===id);
  if(!ev)return;
  const show=document.getElementById("edit-show").value.trim();
  const nombre=document.getElementById("edit-nombre").value.trim();
  const fecha=document.getElementById("edit-fecha").value;
  if(!show||!nombre||!fecha){toast("Completá los campos obligatorios","warn");return;}

  ev.show=show; ev.nombre=nombre; ev.fecha=fecha;
  ev.hora=document.getElementById("edit-hora").value;
  ev.cantidad=parseInt(document.getElementById("edit-cant").value)||2;
  ev.entradasXGan=parseInt(document.getElementById("edit-exg").value)||1;
  ev.venue=document.getElementById("edit-venue").value.trim();
  ev.formUrl=document.getElementById("edit-form").value.trim();
  ev.antiRep=document.getElementById("edit-anti").checked;
  S.ganadores.filter(g=>g.evId===id).forEach(g=>{g.evNombre=nombre;});
  S.eventos.sort((a,b)=>b.fecha.localeCompare(a.fecha));
  const slug2=ev.id.toUpperCase().replace(/[^A-Z0-9]/g,"-").replace(/-+/g,"-").replace(/^-|-$/g,"");
  const closeAt=document.getElementById("edit-closeat").value;
  S.schedules[slug2]=closeAt;
  save();

  // Sincronizar evento actualizado con la nube — await con expectedVersion
  const payload={...ev,expectedVersion:ev._versionAtOpen||""};
  const res=await api({action:"upsertShow",show:payload});
  if(res&&res.error==="conflict"){
    // No revertimos: los edits del usuario quedan en ev y en el form.
    // El modal le da las 3 opciones (Ver cambios / Pisar igual / Cancelar).
    // Si elige Cancelar, ev queda divergente y el sistema se recupera por el siguiente poll o save.
    abrirModalConflicto(ev,res.currentShow,res.currentVersion);
    return;
  }
  if(res&&res.ok&&res.newVersion){
    ev._version=res.newVersion;
    ev._versionAtOpen=res.newVersion;
    save();
  }
  api({action:"setShowSchedule",showId:slug2,closeAt:closeAt}).catch(()=>{});
  closeModal("modal-edit");
  toast("Evento actualizado");
  updateTopbar();
  renderDash();
}
```

Cambios clave respecto a la versión actual:
1. Función ahora es `async`.
2. `payload` incluye `expectedVersion: ev._versionAtOpen||""`.
3. `await` la respuesta del `upsertShow` (antes era fire-and-forget).
4. Si `error==="conflict"` → llama `abrirModalConflicto` (Task 7) y `return` antes de cerrar el modal. **No revierte los cambios locales** — quedan en `ev` y el usuario decide qué hacer desde el modal.
5. Si OK → actualiza `_version` y `_versionAtOpen` con el `newVersion` del servidor.
6. `setShowSchedule` queda fire-and-forget (no es parte de la concurrencia de eventos).

Nota: `abrirModalConflicto` se define en la próxima task. Mantener el orden: Task 7 inmediatamente después de esta.

- [ ] **Step 2: Test happy path (no conflict)**

Recargar admin. Abrir un evento. Cambiar `cantidad` a un valor distinto. Guardar.

Expected:
- Toast "Evento actualizado".
- Modal se cierra.
- En consola: `S.eventos.find(e=>e.id==='<id>')._version` muestra un timestamp más nuevo que antes.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: guardarEvento espera respuesta y manda expectedVersion"
```

---

## Task 7: Frontend — modal de conflicto al guardar

**Files:**
- Modify: `index.html` — agregar CSS dentro del `<style>` existente, agregar HTML antes del `<div class="toast" id="toast">`, agregar funciones JS.

- [ ] **Step 1: Agregar HTML del modal**

Estructura del proyecto: el wrapper exterior usa `class="mbg"` (modal background) con `id` del modal; el `classList.add("show")` se aplica al `mbg`; adentro va un `<div class="modal">`. Botones: `btn bp` (primary), `btn bo` (outlined), `btn br` (rojo/destructivo), sufijo `bsm` para small.

Buscar en `index.html` el bloque `<div class="toast" id="toast">` (alrededor de la línea 1190). Justo ANTES de ese div, insertar:

```html
<div class="mbg" id="modal-conflict">
  <div class="modal" style="width:480px">
    <div class="modal-t">Cambio detectado</div>
    <p style="font-size:13px;color:var(--white);margin-bottom:10px">
      Este evento se modificó en la nube mientras lo editabas.
    </p>
    <p style="font-size:12px;color:var(--gray2);margin-bottom:16px">
      Última actualización remota: <span id="conflict-version">—</span>
    </p>
    <p style="font-size:12px;color:var(--gray2);margin-bottom:18px">
      Para no pisar el trabajo del otro usuario, elegí qué hacer:
    </p>
    <div class="modal-acts">
      <button class="btn bo" onclick="cerrarModalConflicto()">Cancelar</button>
      <button class="btn br bsm" onclick="conflictoPisarIgual()">Pisar igual</button>
      <button class="btn bp" onclick="conflictoVerCambios()">Ver cambios primero</button>
    </div>
  </div>
</div>
```

`abrirModalConflicto` y `cerrarModalConflicto` (Step 2) usan `classList.add("show")` / `.remove("show")` sobre `#modal-conflict` (el `mbg`), igual que el resto de los modales del proyecto. Eso es lo que `closeModal()` ya hace internamente.

- [ ] **Step 2: Agregar las funciones JS**

Buscar la función `guardarEvento` editada en Task 6. Inmediatamente DESPUÉS de su `}` de cierre, agregar:

```javascript
// ─── MODAL DE CONFLICTO ─────────────────────────────────────
let _conflictState=null; // {evId, currentShow, currentVersion}

function abrirModalConflicto(ev,currentShow,currentVersion){
  _conflictState={evId:ev.id,currentShow:currentShow,currentVersion:currentVersion};
  document.getElementById("conflict-version").textContent=currentVersion||"—";
  document.getElementById("modal-conflict").classList.add("show");
}

function cerrarModalConflicto(){
  document.getElementById("modal-conflict").classList.remove("show");
  _conflictState=null;
}

function conflictoVerCambios(){
  if(!_conflictState)return;
  const ev=S.eventos.find(e=>e.id===_conflictState.evId);
  if(!ev){cerrarModalConflicto();return;}
  // Sobrescribir el evento local con currentShow del cloud
  const cs=_conflictState.currentShow||{};
  ev.show=cs.show||ev.show;
  ev.nombre=cs.nombre||ev.nombre;
  ev.fecha=cs.fecha||ev.fecha;
  ev.hora=cs.hora||ev.hora;
  ev.venue=cs.venue||ev.venue;
  ev.cantidad=cs.cantidad||ev.cantidad;
  ev.entradasXGan=cs.entradasXGan||ev.entradasXGan;
  ev.formUrl=cs.formUrl||ev.formUrl;
  ev.antiRep=cs.antiRep!==false;
  ev._version=_conflictState.currentVersion||ev._version;
  ev._versionAtOpen=ev._version;
  save();
  // Recargar el editor con los datos nuevos
  cerrarModalConflicto();
  closeModal("modal-edit");
  abrirEdit(ev.id);
  renderDash();
  toast("Editor recargado con la versión nueva");
}

async function conflictoPisarIgual(){
  if(!_conflictState)return;
  const ev=S.eventos.find(e=>e.id===_conflictState.evId);
  if(!ev){cerrarModalConflicto();return;}
  // ev ya tiene los edits del usuario (guardarEvento no los revierte).
  // Re-llamar upsertShow SIN expectedVersion → pisa el cloud con los edits locales.
  const res=await api({action:"upsertShow",show:ev});
  if(res&&res.ok&&res.newVersion){
    ev._version=res.newVersion;
    ev._versionAtOpen=res.newVersion;
    save();
  }
  cerrarModalConflicto();
  closeModal("modal-edit");
  toast("Evento guardado pisando los cambios remotos");
  updateTopbar();
  renderDash();
}
```

- [ ] **Step 3: Probar con 2 browsers (o 1 browser + 1 incógnito)**

1. Abrir admin en browser A y browser B (uno puede ser incógnito).
2. En A: abrir editor del mismo evento → cambiar `hora` → NO guardar todavía.
3. En B: abrir editor del mismo evento → cambiar `cantidad` → guardar (debería OK).
4. En A: presionar Guardar.

Expected:
- En A, aparece el modal "Cambio detectado" con la versión nueva listada.
- Los cambios locales de A (la hora cambiada) NO se revierten: `S.eventos.find(e=>e.id==='<id>').hora` muestra la hora que A escribió.
- El form de edición sigue visible debajo del modal de conflicto con los valores tipeados por A.
- "Ver cambios primero" → editor se recarga con los datos de B (cantidad actualizada visible), descartando los edits de A.
- "Pisar igual" → guarda con la hora de A, descartando la cantidad de B (que A nunca vio). `S.eventos[ev]._version` se actualiza al `newVersion` devuelto.
- "Cancelar" → solo cierra el modal de conflicto. A sigue con sus edits visibles en el form. Si presiona Guardar de nuevo, vuelve a saltar el modal de conflicto.

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat: modal de conflicto al guardar evento pisado"
```

---

## Task 8: Frontend — función `pollShowsLight` (sin wirear interval todavía)

**Files:**
- Modify: `index.html` — agregar función nueva después de `syncShowsFromCloud`.

- [ ] **Step 1: Implementar `pollShowsLight`**

Buscar la función `syncShowsFromCloud` (~línea 3798). Inmediatamente DESPUÉS de su `}` de cierre, agregar:

```javascript
// ─── POLLER SILENCIOSO ──────────────────────────────────────
let _pollInFlight=false;

async function pollShowsLight(){
  if(document.visibilityState!=="visible")return;
  if(_pollInFlight)return;
  _pollInFlight=true;
  try{
    const res=await api({action:"getShowsCloud"});
    if(!res||!res.ok||!Array.isArray(res.shows)){return;}

    const cloudById=new Map(res.shows.map(s=>[s.id,s]));
    const editorOpen=document.getElementById("modal-edit").classList.contains("show");
    const editorEvId=editorOpen?document.getElementById("edit-id").value:null;
    const sorteoEvId=S.sorteo&&S.sorteo.evId;
    let huboCambios=false;

    // 1. Cloud-side updates y creates
    for(const cs of res.shows){
      const local=S.eventos.find(e=>e.id===cs.id);
      const cloudVer=cs.actualizadoEn||"";
      if(!local){
        // Evento nuevo en cloud
        S.eventos.push({...cs,_version:cloudVer});
        huboCambios=true;
        continue;
      }
      if(cloudVer&&cloudVer!==(local._version||"")){
        if(cs.id===editorEvId){
          // Evento abierto en editor → no pisar campos, solo actualizar _version y mostrar banner
          local._version=cloudVer;
          mostrarBannerActualizado(cs);
          huboCambios=true;
        }else if(cs.id===sorteoEvId){
          // Sorteo en curso → ignorar
          continue;
        }else{
          // Reemplazo silencioso
          Object.assign(local,cs);
          local._version=cloudVer;
          huboCambios=true;
        }
      }
    }

    // 2. Cloud-side deletes (soft-delete: existe local pero no en cloud)
    const cloudIds=new Set(res.shows.map(s=>s.id));
    const aBorrar=[];
    for(const ev of S.eventos){
      if(!cloudIds.has(ev.id)){
        if(ev.id===editorEvId){
          mostrarBannerEliminado(ev);
        }else if(ev.id===sorteoEvId){
          // Sorteo en curso → no borrar
        }else{
          aBorrar.push(ev.id);
        }
      }
    }
    if(aBorrar.length){
      S.eventos=S.eventos.filter(e=>!aBorrar.includes(e.id));
      if(S.evActivo&&aBorrar.includes(S.evActivo)){
        S.evActivo=S.eventos.length?S.eventos[0].id:null;
      }
      huboCambios=true;
    }

    if(huboCambios){
      S.eventos.sort((a,b)=>b.fecha.localeCompare(a.fecha));
      save();
      renderDash();
      updateTopbar();
    }
  }catch(e){
    console.warn("pollShowsLight error:",e);
  }finally{
    _pollInFlight=false;
  }
}

// Stubs — implementados en Tasks 9 y 10
function mostrarBannerActualizado(cloudShow){}
function mostrarBannerEliminado(localEv){}
```

Notas:
- Los stubs `mostrarBannerActualizado` y `mostrarBannerEliminado` se completan en Tasks 9 y 10. Los dejamos como no-ops para que el poller pueda funcionar y testear el merge sin romper.
- El interval se wirea en Task 11. Por ahora la función solo es invocable manualmente desde consola.

- [ ] **Step 2: Probar manualmente desde consola**

Recargar admin en browser A. Esperar 2 segundos. En consola de A:

```javascript
await pollShowsLight();
```

Expected: no errores. `S.eventos` queda igual (no había cambios remotos).

Test con cambio remoto:
1. En browser B (no incógnito, sesión RRHH), editar `cantidad` de un evento y guardar.
2. En A consola: `await pollShowsLight();`
3. En A consola: ver el evento → la cantidad cambió, `_version` cambió.

Expected: el dashboard se re-renderizó con la cantidad nueva.

Test de delete:
1. En B, eliminar un evento (asegurándose de NO tener ese evento abierto en A).
2. En A: `await pollShowsLight();`

Expected: el evento desapareció de `S.eventos` y del dashboard.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: pollShowsLight para sync silencioso de eventos"
```

---

## Task 9: Frontend — banner Estado A (evento actualizado)

**Files:**
- Modify: `index.html` — CSS dentro del `<style>`, HTML dentro de `modal-edit`, reemplazar el stub `mostrarBannerActualizado`.

- [ ] **Step 1: Agregar CSS del banner**

Buscar en `index.html` el bloque `.toast{` dentro del `<style>` (cerca de la línea 203). Justo DESPUÉS de las reglas `.toast`, agregar:

```css
/* BANNER DE CAMBIOS REMOTOS EN EDITOR */
.ev-banner{display:none;background:var(--amber-bg);border:1px solid var(--amber);border-radius:var(--rsm);padding:10px 12px;margin-bottom:14px;font-size:12px;color:var(--white);line-height:1.45}
.ev-banner.on{display:block}
.ev-banner.delete{background:var(--red-bg);border-color:var(--red)}
.ev-banner-msg{margin-bottom:8px}
.ev-banner-actions{display:flex;gap:8px;flex-wrap:wrap}
.ev-banner-btn{font-size:11px;padding:5px 10px;border:1px solid var(--line2);border-radius:var(--rsm);background:transparent;color:var(--white);cursor:pointer;font-family:var(--fb);font-weight:600;transition:all .15s}
.ev-banner-btn:hover{background:rgba(255,255,255,0.08)}
.ev-banner-btn.pri{background:var(--amber);color:var(--bg);border-color:var(--amber)}
.ev-banner.delete .ev-banner-btn.pri{background:var(--red);border-color:var(--red);color:var(--white)}
```

- [ ] **Step 2: Agregar el HTML del banner dentro del `modal-edit`**

Buscar en `index.html` el bloque `<div class="mbg" id="modal-edit">` (línea ~1109). Adentro está `<div class="modal" style="width:520px">` y, primer hijo, `<div class="modal-t">Editar evento</div>` (línea ~1111). Insertar el banner JUSTO DESPUÉS del `modal-t` y ANTES del `<input type="hidden" id="edit-id">`:

```html
<div class="ev-banner" id="edit-banner">
  <div class="ev-banner-msg" id="edit-banner-msg">—</div>
  <div class="ev-banner-actions" id="edit-banner-actions"></div>
</div>
```

- [ ] **Step 3: Reemplazar el stub `mostrarBannerActualizado` con implementación real**

Buscar `function mostrarBannerActualizado(cloudShow){}` (agregado en Task 8) y reemplazar por:

```javascript
let _bannerState=null; // {tipo:'updated'|'deleted', cloudShow|localEv}

function mostrarBannerActualizado(cloudShow){
  _bannerState={tipo:"updated",cloudShow:cloudShow};
  const banner=document.getElementById("edit-banner");
  banner.classList.remove("delete");
  banner.classList.add("on");
  document.getElementById("edit-banner-msg").textContent=
    "Otra persona actualizó este evento. Tu edición está sobre una versión vieja.";
  document.getElementById("edit-banner-actions").innerHTML=
    '<button class="ev-banner-btn pri" onclick="bannerVerCambios()">Ver cambios</button>'+
    '<button class="ev-banner-btn" onclick="bannerSeguirEditando()">Seguir editando</button>';
}

function ocultarBanner(){
  const banner=document.getElementById("edit-banner");
  if(banner){banner.classList.remove("on","delete");}
  _bannerState=null;
}

function bannerVerCambios(){
  if(!_bannerState||_bannerState.tipo!=="updated")return;
  const cs=_bannerState.cloudShow;
  const ev=S.eventos.find(e=>e.id===cs.id);
  if(!ev){ocultarBanner();return;}
  // Pisar campos del evento local con cloud
  Object.assign(ev,cs,{_version:cs.actualizadoEn||ev._version});
  ev._versionAtOpen=ev._version;
  save();
  // Re-cargar los inputs del editor
  document.getElementById("edit-show").value=ev.show||"";
  document.getElementById("edit-nombre").value=ev.nombre||"";
  document.getElementById("edit-fecha").value=ev.fecha||"";
  document.getElementById("edit-hora").value=ev.hora||"";
  document.getElementById("edit-cant").value=ev.cantidad||2;
  document.getElementById("edit-exg").value=ev.entradasXGan||1;
  document.getElementById("edit-venue").value=ev.venue||"";
  document.getElementById("edit-form").value=ev.formUrl||"";
  document.getElementById("edit-anti").checked=ev.antiRep!==false;
  ocultarBanner();
  renderDash();
  toast("Editor actualizado con la versión nueva");
}

function bannerSeguirEditando(){
  // _versionAtOpen NO se toca → al guardar va a saltar el modal de conflicto
  ocultarBanner();
}
```

También, dentro de la función `abrirEdit` (Task 5), agregar al final, antes del `classList.add("show")`:

```javascript
  ocultarBanner();
```

Esto asegura que cada vez que se abre el editor, el banner empieza oculto.

- [ ] **Step 4: Probar con 2 browsers**

1. A y B abren el editor del mismo evento (paso secuencial: A primero, B después).
2. En B: cambiar `cantidad`, guardar.
3. En A (editor abierto): consola → `await pollShowsLight();`

Expected en A:
- Aparece banner amber arriba del form: *"Otra persona actualizó este evento..."*
- Los inputs del form no cambian solos.
- Botón "Ver cambios" → los inputs se recargan con el valor nuevo de `cantidad`, banner desaparece.
- Botón "Seguir editando" → banner desaparece, inputs siguen como estaban; si presionás Guardar después, salta el modal de conflicto (Task 7).

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat: banner contextual en editor cuando otra persona actualiza el evento"
```

---

## Task 10: Frontend — banner Estado B (evento eliminado)

**Files:**
- Modify: `index.html` — reemplazar el stub `mostrarBannerEliminado` con implementación real.

- [ ] **Step 1: Implementar `mostrarBannerEliminado`**

Buscar `function mostrarBannerEliminado(localEv){}` (agregado en Task 8) y reemplazar por:

```javascript
function mostrarBannerEliminado(localEv){
  _bannerState={tipo:"deleted",localEv:localEv};
  const banner=document.getElementById("edit-banner");
  banner.classList.add("on","delete");
  document.getElementById("edit-banner-msg").textContent=
    "Otra persona eliminó este evento. Si guardás se va a restaurar.";
  document.getElementById("edit-banner-actions").innerHTML=
    '<button class="ev-banner-btn pri" onclick="bannerRestaurar()">Restaurar y guardar mis cambios</button>'+
    '<button class="ev-banner-btn" onclick="bannerDescartar()">Descartar</button>';
}

async function bannerRestaurar(){
  if(!_bannerState||_bannerState.tipo!=="deleted")return;
  const ev=_bannerState.localEv;
  // Si el evento ya no está en S.eventos (porque el poller lo borró antes), re-pushear
  if(!S.eventos.find(e=>e.id===ev.id)){
    S.eventos.push(ev);
    S.eventos.sort((a,b)=>b.fecha.localeCompare(a.fecha));
  }
  // Re-aplicar los valores del form al evento local (lo que el usuario tenía en pantalla)
  const formEv=S.eventos.find(e=>e.id===ev.id);
  formEv.show=document.getElementById("edit-show").value.trim()||formEv.show;
  formEv.nombre=document.getElementById("edit-nombre").value.trim()||formEv.nombre;
  formEv.fecha=document.getElementById("edit-fecha").value||formEv.fecha;
  formEv.hora=document.getElementById("edit-hora").value;
  formEv.cantidad=parseInt(document.getElementById("edit-cant").value)||formEv.cantidad;
  formEv.entradasXGan=parseInt(document.getElementById("edit-exg").value)||formEv.entradasXGan;
  formEv.venue=document.getElementById("edit-venue").value.trim();
  formEv.formUrl=document.getElementById("edit-form").value.trim();
  formEv.antiRep=document.getElementById("edit-anti").checked;
  save();
  // Crear de vuelta en el cloud — sin expectedVersion (es un restore consciente)
  const res=await api({action:"upsertShow",show:formEv});
  if(res&&res.ok&&res.newVersion){
    formEv._version=res.newVersion;
    formEv._versionAtOpen=res.newVersion;
    save();
  }
  ocultarBanner();
  closeModal("modal-edit");
  toast("Evento restaurado");
  updateTopbar();
  renderDash();
}

function bannerDescartar(){
  if(!_bannerState||_bannerState.tipo!=="deleted")return;
  const id=_bannerState.localEv.id;
  S.eventos=S.eventos.filter(e=>e.id!==id);
  if(S.evActivo===id){S.evActivo=S.eventos.length?S.eventos[0].id:null;}
  save();
  ocultarBanner();
  closeModal("modal-edit");
  toast("Evento descartado");
  updateTopbar();
  renderDash();
}
```

Nota: en el `pollShowsLight` (Task 8), el delete-loop deja al evento en `S.eventos` cuando está abierto en el editor (no entra en `aBorrar`). Eso es intencional: `mostrarBannerEliminado(ev)` recibe el `ev` local todavía vivo. Si el usuario no toca el banner y mientras tanto cierra el editor (sin restaurar ni descartar), el evento queda zombi en local hasta el próximo poll. Eso es OK — el próximo poll lo va a barrer porque `modal-edit` ya no estará abierto.

- [ ] **Step 2: Probar con 2 browsers**

1. A y B abren editor del mismo evento.
2. En B: eliminar el evento (botón eliminar dentro del modal).
3. En A (editor todavía abierto): consola → `await pollShowsLight();`

Expected en A:
- Aparece banner rojo arriba del form: *"Otra persona eliminó este evento..."*
- "Restaurar y guardar mis cambios" → el evento aparece de vuelta en A y B (después de que B haga un poll o recargue).
- "Descartar" → cierra el editor en A, el evento sigue eliminado en ambos.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: banner de evento eliminado con restaurar/descartar"
```

---

## Task 11: Wirear el interval del poller

**Files:**
- Modify: `index.html:3825-3833` (función `startApp`)

- [ ] **Step 1: Agregar `setInterval` en `startApp`**

Reemplazar la función `startApp` por:

```javascript
function startApp(){
  load();
  renderDash();
  updateTopbar();
  // Cargar shows desde la nube (~1s de delay para no bloquear render inicial)
  setTimeout(function(){syncShowsFromCloud();},800);
  // Cargar tracking de ganadores (~1.5s)
  setTimeout(function(){importarDesdeTracking(true);},1500);
  // Polling silencioso cada 60s para detectar cambios remotos en eventos
  setInterval(pollShowsLight,60000);
}
```

Única línea agregada: `setInterval(pollShowsLight,60000);`.

- [ ] **Step 2: Smoke test del poller en runtime**

Recargar el admin. Abrir DevTools → Network → filtrar por `script.google.com`. Dejar la pestaña visible y esperar 65 segundos.

Expected: aparece una request `POST` a `script.google.com` aproximadamente a los 60 s del inicio, con `getShowsCloud` en el body. Response 200.

Esperar otros 60 s. Expected: otra request más.

Cambiar a otra pestaña del browser y esperar 90 s. Volver al admin. Expected: durante esos 90 s no hubo requests (gate de visibilityState funcionando).

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: wirear poller de eventos cada 60s en startApp"
```

---

## Task 12: Verificación final — checklist completo del spec

**Files:** ninguno (testing manual)

- [ ] **Step 1: Ejecutar las 11 verificaciones del spec**

Tener 2 browsers abiertos (A y B). Recargar ambos antes de empezar.

Test 1 — Crear evento desde A se ve en B:
- A: crear evento nuevo "TEST CONCURRENCIA" con fecha futura. Apretar Crear.
- B: esperar 60 s. Apretar `await pollShowsLight();` en consola si querés acelerar.
- Expected: evento aparece en dashboard de B sin recargar.

Test 2 — Editar cantidad desde A se ve en B:
- A: abrir editor de "TEST CONCURRENCIA", cambiar cantidad a 5, guardar.
- B: esperar el siguiente poll.
- Expected: cantidad 5 visible en dashboard de B.

Test 3 — Eliminar desde A desaparece en B:
- A: abrir editor, eliminar.
- B: esperar el siguiente poll.
- Expected: evento desaparece del dashboard de B.

Test 4 — Conflict modal funciona:
- (Crear evento "TEST2" desde A.)
- A y B abren editor de "TEST2". A guarda con cantidad 3. B guarda con cantidad 7.
- Expected: B ve modal "Cambio detectado". A queda con cantidad 3.

Test 5 — "Ver cambios primero" recarga form sin error:
- En el escenario de Test 4, B aprieta "Ver cambios primero".
- Expected: editor de B se recarga con cantidad 3. No error en consola.

Test 6 — "Pisar igual" sobrescribe:
- Repetir Test 4. B aprieta "Pisar igual".
- Expected: guarda con cantidad 7. A en el siguiente poll ve banner amber.

Test 7 — Banner de eliminado:
- A y B abren editor de "TEST2". B elimina.
- A: consola → `await pollShowsLight();`
- Expected: banner rojo en A.

Test 8 — Restaurar funciona:
- En el escenario de Test 7, A aprieta "Restaurar y guardar mis cambios".
- Expected: evento vuelve a aparecer. En B, después del próximo poll, también aparece.

Test 9 — Pestaña en background:
- Cambiar de pestaña en A. En B, modificar el evento. Esperar 90 s.
- Volver a la pestaña de A. Expected: en ≤60 s el cambio aparece en A.

Test 10 — Red intermitente:
- En A, abrir DevTools → Network → setear "Offline". Esperar 2 minutos. Volver "Online".
- Expected: ningún toast de error apareció. El próximo poll funciona.

Test 11 — Sorteo en curso no se sincroniza:
- A inicia sorteo en evento X (apretar Sortear, sin confirmar).
- B modifica el evento X (cantidad).
- A: `await pollShowsLight();`
- Expected: ni banner ni cambio visible en el evento X de A. (Verificable en consola: `S.eventos.find(e=>e.id==='X').cantidad` no cambió.)

- [ ] **Step 2: Push a `main` para deployar**

Si todos los tests pasaron, push:

```bash
git push origin main
```

GitHub Pages deploya automáticamente. Esperar ~1 minuto. Recargar la app desde la URL pública y repetir Test 4 (el más representativo) en producción.

- [ ] **Step 3: Avisar a la otra persona de RRHH**

Mensaje sugerido por Slack/Teams/WhatsApp interno:

> "Hola — actualicé el sistema de sorteos: ahora cuando los dos editamos el mismo evento al mismo tiempo el sistema avisa en vez de pisar cambios. Si te aparece un cartel amarillo arriba del editor diciendo que actualicé algo, o un modal al guardar, leelo antes de seguir. Cualquier cosa rara, gritame."

---

## Notas finales

- **No hay rollback automático del deploy del Apps Script**. Si la nueva versión rompe algo, en Apps Script: `Deploy → Manage deployments → Edit → Version → seleccionar una anterior → Deploy`. La URL no cambia.
- **localStorage no se migra**: navegadores con la versión vieja siguen funcionando, solo que sus eventos van a tener `_version` undefined hasta el primer sync, que es lo que va a poblar el campo. No hay migration step.
- **`creadoEn` vs `actualizadoEn`**: el cloud guarda ambos. El frontend hoy solo conocía `creadoEn`. Tras este plan, `actualizadoEn` viaja como `_version` (renombrado intencional para que el campo del frontend describa su rol, no su origen).
