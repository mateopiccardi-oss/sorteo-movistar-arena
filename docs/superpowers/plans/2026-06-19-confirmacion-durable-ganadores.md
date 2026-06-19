# Confirmación durable de ganadores — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que un sorteo no pueda darse por "confirmado" si los ganadores no quedaron escritos en Sheets, y que un ganador no sincronizado nunca se pierda al reconstruir el estado local desde Sheets.

**Architecture:** Se agrega una bandera por ganador `_synced` (`false` al confirmar, `true` solo cuando `guardarEnSheets` responde OK). La reconstrucción desde Sheets (`importarTracking`) preserva los ganadores `_synced:false` en vez de descartarlos. La UI muestra un banner persistente "sin sincronizar — Reintentar" en las pantallas de Ganadores y Envío mientras existan ganadores no sincronizados del evento activo; el botón reintenta la escritura. El enfoque es **optimista + durable**: se mantiene el reveal inmediato actual, pero el fallo deja de ser un toast efímero y pasa a ser un estado visible y recuperable.

**Tech Stack:** HTML/JS vanilla (un solo `index.html`), backend Google Apps Script (sin cambios en este plan). Sin framework de tests — la verificación es manual en el navegador simulando un fallo de sync, más revisión estática del diff.

**Spec / diagnóstico:** memoria `project-bug-envio-angela-leiva` (causa raíz confirmada 2026-06-19). Este plan cubre los **defectos 1 (confirmación no durable)** y **4 (estado local frágil multi-user)**. Los defectos 2 (ID = nombre inestable) y 3 ("0 pendientes" = éxito mudo) quedan para un segundo plan.

## Global Constraints

- Un solo archivo de UI: `index.html`. No hay build ni bundler; los cambios son JS inline dentro de `<script>`.
- No romper el flujo optimista existente: el reveal de ganadores debe seguir apareciendo inmediato tras confirmar.
- `_synced` es un campo nuevo y opcional: los ganadores viejos en localStorage no lo tienen → tratar `_synced===false` (estricto) como "no sincronizado" y `undefined`/`true` como "sincronizado/legacy", para no marcar histórico previo como pendiente de sync.
- No tocar el backend (`sorteo_script.gs`) en este plan.
- Mantener el estilo de código del archivo (sin semicolons opcionales nuevos, comillas dobles, funciones `function` o arrow según contexto vecino).

---

### Task 1: Marcar ganadores como `_synced` al confirmar

Hace que el modelo de datos distinga ganadores escritos en Sheets de los que solo viven en el navegador. Es la base de las tareas siguientes.

**Files:**
- Modify: `index.html` — `confirmarGanadores()` (`index.html:3302-3347`)

**Interfaces:**
- Produces: cada objeto en `S.ganadores` puede tener `_synced: boolean`. `false` = escrito en local pero NO confirmado por Sheets. `true` = confirmado por `guardarEnSheets`. Ausente = dato legacy, se trata como sincronizado.

- [ ] **Step 1: Reproducir el fallo (verificación manual previa)**

Abrir `index.html` en el navegador con DevTools. En la consola, stubear la API para forzar el fallo de guardado:

```js
const _origApi = window.api;
window.api = async (p) => { if (p.action === "guardarGanadores") throw new Error("SIMULADO: sin red"); return _origApi(p); };
```

Hacer un sorteo de prueba y confirmar. Observar: aparece el reveal con los ganadores y un toast efímero "⚠ Error al sincronizar Sheets" que se va. Recargar la página → los ganadores **siguen** en la lista local pero **no** están en Sheets. Este es el bug. Restaurar con `window.api=_origApi`.

- [ ] **Step 2: Setear `_synced:false` al empujar los ganadores**

En `confirmarGanadores()`, reemplazar el bloque del push (`index.html:3311-3315`):

```javascript
  lista.forEach(g=>S.ganadores.push({
    ...g, id:"g"+Date.now()+"_"+Math.floor(Math.random()*99999), evId, evNombre:ev.nombre,
    venue:ev.venue, fecha:ev.fecha, hora:ev.hora,
    estado:"pendiente", fechaGano:new Date().toISOString()
  }));
```

por:

```javascript
  lista.forEach(g=>S.ganadores.push({
    ...g, id:"g"+Date.now()+"_"+Math.floor(Math.random()*99999), evId, evNombre:ev.nombre,
    venue:ev.venue, fecha:ev.fecha, hora:ev.hora,
    estado:"pendiente", fechaGano:new Date().toISOString(),
    _synced:false
  }));
```

- [ ] **Step 3: Marcar `_synced:true` solo en el éxito de la sync**

En el branch de éxito (`index.html:3334-3346`), dentro del `else`, agregar el marcado de `_synced:true` junto al `_historico=true` existente. Reemplazar:

```javascript
    lista.forEach(g=>{
      const localG=S.ganadores.find(x=>x.evId===evId&&x.nombre===g.nombre&&!x._historico);
      if(localG)localG._historico=true;
    });
```

por:

```javascript
    lista.forEach(g=>{
      const localG=S.ganadores.find(x=>x.evId===evId&&x.nombre===g.nombre&&x._synced===false);
      if(localG){localG._historico=true;localG._synced=true;}
    });
```

(Se cambia el predicado `!x._historico` por `x._synced===false` para apuntar exactamente a los recién confirmados de esta tanda.)

- [ ] **Step 4: Verificar**

Recargar la página (sin stub). Hacer un sorteo y confirmar con red OK. En consola:

```js
S.ganadores.filter(g=>g._synced===false)
```

Expected: `[]` tras una sync exitosa (todos quedaron `_synced:true`). Repetir con el stub de fallo del Step 1 → los ganadores quedan con `_synced:false`.

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat(ganadores): marcar _synced en confirmarGanadores para distinguir lo escrito en Sheets"
```

---

### Task 2: Preservar ganadores no sincronizados al reconstruir desde Sheets

Cierra el vector de pérdida de datos: hoy `importarTracking` hace `S.ganadores=[]` y reconstruye desde Sheets, borrando para siempre cualquier ganador cuya sync falló.

**Files:**
- Modify: `index.html` — bloque de reconstrucción en `importarTracking` (`index.html:2469-2472`)

**Interfaces:**
- Consumes: `_synced` de Task 1.
- Produces: tras la reconstrucción, `S.ganadores` contiene los ganadores de Sheets **más** los locales `_synced:false`.

- [ ] **Step 1: Reproducir la pérdida (verificación manual previa)**

Con el estado del Step 4 de Task 1 (un evento con ganadores `_synced:false` por fallo simulado), ejecutar la importación desde Sheets (el botón "Importar tracking" o `importarTracking(true)` en consola). Observar: los ganadores `_synced:false` **desaparecen** de `S.ganadores`. Esto es la pérdida.

- [ ] **Step 2: Preservar los no sincronizados antes del rebuild**

Reemplazar (`index.html:2469-2472`):

```javascript
    // Rebuild ganadores entirely from Sheets. Non-historico local ganadores are safe to
    // discard because confirmarGanadores() marks them _historico right after a successful
    // Sheets sync; any that remain non-historico are stale/garbage data.
    S.ganadores=[];
    S.eventos=S.eventos.filter(function(e){return !e._fromTracking;});
```

por:

```javascript
    // Rebuild ganadores from Sheets, PERO preservando los locales no sincronizados:
    // un ganador con _synced===false nunca llegó a Sheets, así que el rebuild lo borraría
    // para siempre (este fue el bug de Angela Leiva). Los conservamos para que el banner
    // de "sin sincronizar" siga ofreciendo reintentar.
    const _noSync=S.ganadores.filter(function(g){return g._synced===false;});
    S.ganadores=_noSync.slice();
    S.eventos=S.eventos.filter(function(e){return !e._fromTracking;});
```

- [ ] **Step 3: Verificar**

Repetir el Step 1. En consola tras importar:

```js
S.ganadores.filter(g=>g._synced===false).length
```

Expected: igual a la cantidad de ganadores no sincronizados que había antes de importar (ya **no** se pierden). Verificar además que los ganadores que sí estaban en Sheets se reconstruyeron normalmente (lista total razonable, sin duplicados de los `_synced:false`).

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "fix(ganadores): no descartar ganadores _synced:false al reconstruir desde Sheets"
```

---

### Task 3: Banner persistente "sin sincronizar" + reintento

Reemplaza el toast efímero por un estado visible y accionable. Mientras el evento activo tenga ganadores `_synced:false`, se muestra un banner con botón "Reintentar sincronización" en las pantallas de Ganadores y Envío.

**Files:**
- Modify: `index.html` — agregar `renderSyncBanner()` y `reintentarSync()` (cerca de `guardarEnSheets`, ~`index.html:3597`); llamar al render desde `renderGanConf()` (`index.html:3599`) y `renderMails()` (`index.html:3684`); ajustar el branch de error de `confirmarGanadores()` (`index.html:3331-3333`)

**Interfaces:**
- Consumes: `_synced` (Task 1), `S.ganadores`, `S.evActivo`, `S.eventos`, `guardarEnSheets`, `trackingEnSheets`, `toast`, `save`.
- Produces: `renderSyncBanner(containerId)` inyecta/actualiza un banner en el contenedor dado. `reintentarSync(evId)` reintenta la escritura de los `_synced:false` del evento y, si OK, los marca `_synced:true`.

- [ ] **Step 1: Agregar `renderSyncBanner` y `reintentarSync`**

Insertar estas dos funciones inmediatamente después de `trackingEnSheets` (tras `index.html:3597`):

```javascript
// Banner persistente: ganadores confirmados localmente pero NO escritos en Sheets.
// Reemplaza al toast efímero — el riesgo queda visible hasta resolverse.
function unsyncedDelEvento(evId){
  return S.ganadores.filter(g=>g.evId===evId&&g._synced===false);
}
function renderSyncBanner(containerId){
  const cont=document.getElementById(containerId);
  if(!cont)return;
  const evId=S.evActivo;
  const pend=evId?unsyncedDelEvento(evId):[];
  let banner=cont.querySelector(".sync-banner");
  if(!pend.length){if(banner)banner.remove();return;}
  if(!banner){
    banner=document.createElement("div");
    banner.className="sync-banner";
    banner.style.cssText="display:flex;align-items:center;gap:10px;padding:11px 14px;margin-bottom:12px;border-radius:10px;background:rgba(255,176,32,0.10);border:1px solid rgba(255,176,32,0.45)";
    cont.prepend(banner);
  }
  banner.innerHTML=
    '<div style="font-size:18px">⚠</div>'+
    '<div style="flex:1;font-size:12px;color:var(--white);line-height:1.4">'+
      '<strong>'+pend.length+' ganador'+(pend.length!==1?'es':'')+' sin sincronizar con Sheets.</strong> '+
      'No se escribieron en la planilla, así que <strong>no se les puede enviar</strong> hasta sincronizar.'+
    '</div>'+
    '<button class="btn bp bsm" id="btn-reintentar-sync">Reintentar</button>';
  const btn=banner.querySelector("#btn-reintentar-sync");
  if(btn)btn.onclick=()=>reintentarSync(evId);
}
async function reintentarSync(evId){
  const ev=S.eventos.find(e=>e.id===evId);
  const pend=unsyncedDelEvento(evId);
  if(!ev||!pend.length){toast("No hay ganadores sin sincronizar","warn");return;}
  const btn=document.getElementById("btn-reintentar-sync");
  if(btn){btn.disabled=true;btn.textContent="Sincronizando…";}
  const lista=pend.map(g=>({nombre:g.nombre,email:g.email}));
  const resultados=await Promise.allSettled([guardarEnSheets(lista,ev),trackingEnSheets(lista,ev)]);
  const errores=resultados.filter(r=>r.status==="rejected").map(r=>r.reason?.message||"Error desconocido");
  if(errores.length){
    if(btn){btn.disabled=false;btn.textContent="Reintentar";}
    toast("⚠ Sigue fallando la sincronización: "+errores[0],"warn");
    return;
  }
  pend.forEach(g=>{g._synced=true;g._historico=true;});
  save();
  toast("✓ "+pend.length+" ganador(es) sincronizado(s) con Sheets");
  renderGanConf();renderMails();renderDash();
}
```

- [ ] **Step 2: Llamar al banner desde `renderGanConf` y `renderMails`**

En `renderGanConf()` (`index.html:3599`), como primera línea del cuerpo de la función agregar:

```javascript
  renderSyncBanner("gan-conf-list");
```

En `renderMails()` (`index.html:3684`), como primera línea del cuerpo de la función agregar:

```javascript
  renderSyncBanner("mail-ct");
```

(Si `mail-ct` no es un contenedor apto para `prepend`, usar el contenedor de la lista de mails de esa función — ver Step 4 de verificación. El elemento `mail-ct` existe en `index.html:988`.)

- [ ] **Step 3: Hacer persistente el aviso en el branch de error de `confirmarGanadores`**

En `confirmarGanadores()`, reemplazar el branch de error (`index.html:3331-3333`):

```javascript
  if(errores.length){
    toast("⚠ Error al sincronizar Sheets: "+errores[0],"warn");
    console.error("Sheets sync errors:",errores);
  } else {
```

por:

```javascript
  if(errores.length){
    toast("⚠ No se pudo sincronizar con Sheets — quedó pendiente de reintento","warn");
    console.error("Sheets sync errors:",errores);
    renderGanConf();renderMails();
  } else {
```

(Los ganadores ya quedaron `_synced:false` por Task 1; acá solo se fuerza el render del banner para que el aviso sea persistente en vez de efímero.)

- [ ] **Step 4: Verificar (end-to-end)**

Con el stub de fallo (Task 1, Step 1) activo: hacer un sorteo y confirmar. Expected: aparece el banner ámbar "N ganadores sin sincronizar" en la pantalla de Ganadores y persiste tras navegar/recargar (los `_synced:false` están en localStorage). Quitar el stub (`window.api=_origApi`) y apretar "Reintentar". Expected: el banner desaparece, toast "✓ N sincronizado(s)", y en consola `S.ganadores.filter(g=>g._synced===false)` da `[]`. Confirmar también que los ganadores ahora están en la pestaña Ganadores de Sheets.

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat(ganadores): banner persistente de sync + reintento en vez de toast efímero"
```

---

### Task 4: Verificación integral y limpieza de la entrada duplicada

Cierra el caso: prueba el escenario completo de Angela Leiva y limpia el residuo de config.

**Files:**
- Sin cambios de código. Verificación manual + limpieza de datos en la hoja Shows.

- [ ] **Step 1: Escenario completo**

1. Crear un show de prueba, sortear, confirmar con red OK → verificar fila escrita en Ganadores (Sheets) y `_synced:true`.
2. Repetir con red caída (stub) → banner persistente, sin fila en Sheets.
3. Restaurar red, reintentar desde el banner → fila aparece en Sheets, banner desaparece.
4. Forzar `importarTracking` con un evento `_synced:false` pendiente → el ganador NO se pierde (Task 2).

- [ ] **Step 2: Limpiar la entrada duplicada de Angela Leiva**

En la hoja **Shows** ya hay dos filas de Angela Leiva: `ANGELA LEIVA` (activa, cupo 15) y `2026-06-18 ANGELA LEIVA` (Eliminado=1, cupo 30). Confirmar con Mateo cuál es la correcta (probablemente la activa cupo 15) y dejar solo esa. La eliminada ya tiene `Eliminado=1`, así que no aparece en `getShowsCloud` — no requiere acción salvo confirmar que no quede referencia local en algún navegador. Documentar que el origen del duplicado (defecto 2: `id = nombre`) se aborda en el plan siguiente.

- [ ] **Step 3: Actualizar la memoria del proyecto**

Marcar en `project-bug-envio-angela-leiva` que los defectos 1 y 4 quedaron resueltos por este plan, y dejar 2 y 3 como pendientes para el plan de "clave estable".

---

## Notas para el plan siguiente (fuera de alcance acá)

- **Defecto 2 (ID inestable):** introducir un `slug` estable único por show, generado una sola vez en la creación y usado como `showId` en TODAS las operaciones (`guardarGanadores`, `enviarMails`, `getInscriptos`, `setShowActivo`, `checkPDFs`). Requiere migrar los IDs existentes en las hojas Shows/Ganadores.
- **Defecto 3 ("0 pendientes" = éxito mudo):** que el front (`index.html:3778`) y el backend (`sorteo_script.gs:340`) distingan "todo enviado" de "no hay nada que enviar pero debería haber", y avisen.
