# Entradas por ganador + envío por lotes · Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cantidad de entradas editable por ganador (con default por evento), envío de mails por lotes con reanudación automática sin reusar PDFs, y envío verificado para los dos operadores.

**Architecture:** Backend primero (`sorteo_script.gs`, retro-compatible con el frontend actual), deploy, y después el frontend (`index.html`, single-file vanilla JS). La pestaña **Ganadores** del Sheet gana la columna J "Entradas" y pasa a ser la fuente de verdad de cantidades; el envío corta a los 4 minutos y la app re-llama hasta terminar.

**Tech Stack:** HTML + JavaScript vanilla embebido en `index.html` · Google Apps Script (`sorteo_script.gs`). Sin framework de tests → **verificación por consola del navegador** y por ejecución manual en el editor de Apps Script (método ya usado en este proyecto).

**Spec:** `docs/superpowers/specs/2026-07-05-entradas-por-ganador-envio-lotes-design.md`

**Dos extensiones respecto de la spec** (necesarias para cumplir sus propias pruebas, detectadas al mapear el código):

1. **Endpoint `getGanadoresShow`**: la app reconstruye `S.ganadores` desde el Tracking al recargar (`importarDesdeTracking`, [index.html:2493](../../../index.html)), o sea que `g.entradas` local se pierde. Para la prueba 3 de la spec (persistencia multi-compu) el frontend hidrata las cantidades leyendo la pestaña Ganadores.
2. **Ajuste de B123 al editar entradas**: el total histórico de tickets (celda B123 del Tracking) se incrementa al confirmar; si después se edita la cantidad de un pendiente, `actualizarEntradas` aplica el delta para que el total no quede desfasado.

**Bug preexistente que este plan arregla de paso:** el router ([sorteo_script.gs:107](../../../sorteo_script.gs)) llama `trackingGanadores(body.ganadores, body.showNombre, body.fecha)` **sin** el 4º parámetro `entradasXGan` que el frontend sí manda → B123 siempre suma 1 por ganador aunque el evento dé más entradas.

## Global Constraints

- **NO cambiar el modo de deploy del Apps Script:** queda "Ejecutar como: Yo" + "Cualquier persona". Para publicar cambios: Implementar → **Administrar implementaciones** → ✏️ editar → Versión: **Nueva versión** → Implementar (misma URL). NUNCA "Nueva implementación" (cambia la URL) ni "Usuario que accede" (rompe CORS).
- **Backend retro-compatible antes que frontend:** cada cambio en `.gs` debe funcionar con el frontend viejo (fallbacks a `entradasXGan` del request y a col J vacía).
- La app es **un único `index.html` autocontenido**: sin archivos JS externos, sin frameworks, rutas relativas.
- **No tocar columnas A y B de "Tracking Ganadores"** (tienen fórmulas). La matriz de nombres no cambia (un nombre por ganador, sin repetir).
- Anclar ediciones por **nombre de función**, no por número de línea (las líneas se corren).
- Cantidades válidas: entero **1–10** por ganador.

## Quién ejecuta qué

- **Agente:** edita `index.html` y `sorteo_script.gs` en el repo y hace los commits.
- **Mateo:** pega el `.gs` en el editor de Apps Script, publica la **nueva versión** del deploy existente, y corre las verificaciones en la app real / editor / planilla. Los pasos suyos están marcados "(Mateo)".

## Estructura de archivos

| Archivo | Responsabilidad | Cambio |
|---|---|---|
| `sorteo_script.gs` | Backend Apps Script | `guardarGanadores`, `_enviarMailsLocked`, `checkPDFs`, `trackingGanadores`, router `doPost`; nuevos: `_pdfsUsadosDelShow`, `actualizarEntradas`, `getGanadoresShow` |
| `index.html` | App admin completa | `confirmarGanadores`, `guardarEnSheets`, `trackingEnSheets`, `renderGanConf`, `renderMails`, `_enviarMailsCore`, `mostrarPreflightPDFs`, `drawTickets` (×2), `renderDash`; nuevos: `entradasDe`, `stepEntradas`, `_pushEntradas`, `hydrateEntradas` |

---

## Task 1: Backend — columna J "Entradas" en `guardarGanadores`

**Files:**
- Modify: `sorteo_script.gs` — función `guardarGanadores` (≈259-288)

**Interfaces:**
- Consumes: `body.ganadores` puede traer `entradas` (entero) por ganador; el frontend viejo no lo manda.
- Produces: pestaña "Ganadores" con header J1 = "Entradas"; cada fila nueva tiene col J = entero, o **vacía** si el request no trajo `entradas` (clave para retro-compatibilidad: col J vacía → el envío usa el `entradasXGan` del request).

- [ ] **Step 1 (Agente):** En `guardarGanadores`, reemplazar el bloque de creación de hoja y el `appendRow` de filas:

Reemplazar:

```js
  if (!hoja) {
    hoja = ss.insertSheet("Ganadores");
    hoja.appendRow(["Timestamp", "Show ID", "Show", "Fecha Show", "Venue", "Mail", "Nombre", "Estado Mail", "PDF enviado"]);
    hoja.getRange(1, 1, 1, 9).setFontWeight("bold");
  }
```

por:

```js
  if (!hoja) {
    hoja = ss.insertSheet("Ganadores");
    hoja.appendRow(["Timestamp", "Show ID", "Show", "Fecha Show", "Venue", "Mail", "Nombre", "Estado Mail", "PDF enviado", "Entradas"]);
    hoja.getRange(1, 1, 1, 10).setFontWeight("bold");
  }
  // Migración lazy: hojas creadas antes de la col "Entradas"
  if (!String(hoja.getRange(1, 10).getValue()).trim()) {
    hoja.getRange(1, 10).setValue("Entradas").setFontWeight("bold");
  }
```

y reemplazar el cuerpo del `forEach`:

```js
  ganadores.forEach(g => {
    hoja.appendRow([
      new Date(),
      showId,
      showNombre,
      fecha || "",
      venue || "Movistar Arena",
      g.mail || g.email,
      g.nombre,
      "Pendiente",
      ""
    ]);
  });
```

por:

```js
  ganadores.forEach(g => {
    const ent = parseInt(g.entradas);
    hoja.appendRow([
      new Date(),
      showId,
      showNombre,
      fecha || "",
      venue || "Movistar Arena",
      g.mail || g.email,
      g.nombre,
      "Pendiente",
      "",
      (isNaN(ent) || ent < 1) ? "" : ent
    ]);
  });
```

- [ ] **Step 2 (Agente):** Commit.

```bash
git add sorteo_script.gs
git commit -m "feat(backend): columna Entradas (col J) por fila en pestana Ganadores"
```

---

## Task 2: Backend — helper `_pdfsUsadosDelShow` + `_enviarMailsLocked` por ganador, por lotes, sin reusar PDFs

**Files:**
- Modify: `sorteo_script.gs` — función `_enviarMailsLocked` (≈316-413); helper nuevo arriba de `enviarMails`

**Interfaces:**
- Consumes: col J de "Ganadores" (Task 1); `buscarPDFs(showId, showNombre)` existente (devuelve Files ordenados por nombre).
- Produces: `_pdfsUsadosDelShow(datos, showId)` → objeto `{nombrePdf: true}`. Respuesta de `enviarMails` gana el campo **`restantes`** (entero: pendientes NO intentados en esta ejecución; los que fallaron con error NO cuentan como restantes, para que el loop del frontend no quede infinito). Secuencia por fila: `Enviando` + col I (PDFs) **antes** de `sendEmail`; `Enviado` después; si falla sin enviar → vuelve a `Pendiente` y col I se limpia.

- [ ] **Step 1 (Agente):** Agregar el helper justo antes de `function enviarMails(...)`:

```js
// PDFs ya asignados a CUALQUIER fila del show (col I "PDF enviado").
// Evita que una reanudación vuelva a repartir PDFs ya enviados a otros ganadores.
function _pdfsUsadosDelShow(datos, showId) {
  var usados = {};
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() !== String(showId).trim()) continue;
    String(datos[i][8] || "").split(",").forEach(function(nom) {
      var n = nom.trim();
      if (n) usados[n] = true;
    });
  }
  return usados;
}
```

- [ ] **Step 2 (Agente):** Reemplazar `_enviarMailsLocked` completa por:

```js
function _enviarMailsLocked(showId, entradasXGan) {
  if (!showId) return { ok: false, error: "Show ID requerido" };
  var inicio = Date.now();
  var BUDGET_MS = 240000; // corta limpio a los 4 min (limite real de Apps Script: ~6 min)
  var epgDefault = parseInt(entradasXGan) || 1;

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  const hoja = ss.getSheetByName("Ganadores");
  if (!hoja) return { ok: false, error: "No hay ganadores registrados" };

  const datos = hoja.getDataRange().getValues();
  const ganadores = [];

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() === String(showId).trim() && String(datos[i][7]).trim() === "Pendiente") {
      const entFila = parseInt(datos[i][9]);
      ganadores.push({
        mail: String(datos[i][5]).trim(),
        nombre: String(datos[i][6]).trim(),
        showNombre: String(datos[i][2]).trim(),
        fecha: String(datos[i][3]).trim(),
        venue: String(datos[i][4]).trim(),
        fila: i + 1,
        entradas: (isNaN(entFila) || entFila < 1) ? epgDefault : entFila
      });
    }
  }

  if (!ganadores.length) return { ok: true, enviados: 0, enviadosMails: [], restantes: 0, mensaje: "No hay ganadores pendientes" };

  // Buscar PDFs y descartar los ya asignados en cualquier fila del show
  const usados = _pdfsUsadosDelShow(datos, showId);
  const pdfs = buscarPDFs(showId, ganadores[0].showNombre).filter(function(p) { return !usados[p.getName()]; });

  // Regla estricta: si los PDFs disponibles no alcanzan, no se envía NINGÚN mail
  const necesarios = ganadores.reduce(function(acc, g) { return acc + g.entradas; }, 0);
  if (pdfs.length < necesarios) {
    return {
      ok: false,
      error: "Hay " + pdfs.length + " PDF(s) disponibles en Drive (sin contar los ya enviados) y se necesitan " + necesarios +
             " para " + ganadores.length + " ganador(es) pendiente(s) — no se envió ningún mail."
    };
  }

  const errores = [];
  const enviadosMails = [];
  let enviados = 0;
  let pdfCursor = 0;
  let idx = 0;

  for (idx = 0; idx < ganadores.length; idx++) {
    if (Date.now() - inicio > BUDGET_MS) break; // corta limpio; la app reanuda con otra llamada
    const g = ganadores[idx];
    let sent = false;
    try {
      const asunto = CONFIG.MAIL_ASUNTO
        .replace(/{nombre}/g, g.nombre)
        .replace(/{evento}/g, g.showNombre);

      const cuerpo = CONFIG.MAIL_CUERPO
        .replace(/{nombre}/g, g.nombre)
        .replace(/{evento}/g, g.showNombre)
        .replace(/{venue}/g, g.venue)
        .replace(/{fecha}/g, g.fecha);

      const opts = {
        name: CONFIG.MAIL_REMITENTE,
        from: CONFIG.MAIL_FROM_ALIAS,
        replyTo: CONFIG.MAIL_FROM_ALIAS
      };

      // Cada ganador consume g.entradas PDFs de la lista disponible
      const pdfSlice = pdfs.slice(pdfCursor, pdfCursor + g.entradas);
      const pdfNombres = pdfSlice.map(function(p) { return p.getName(); }).join(", ");
      opts.attachments = pdfSlice.map(function(p) { return p.getAs(MimeType.PDF); });

      // Enviar como HTML para soportar UTF-8 y emojis correctamente
      opts.htmlBody = cuerpo.replace(/\n/g, "<br>");

      // Marcar estado Y PDFs ANTES de enviar: si la ejecución muere acá, un reintento
      // no re-manda este mail (no es Pendiente) ni reusa estos PDFs (figuran en col I)
      hoja.getRange(g.fila, 8).setValue("Enviando");
      hoja.getRange(g.fila, 9).setValue(pdfNombres);
      GmailApp.sendEmail(g.mail, asunto, cuerpo, opts);
      sent = true;
      hoja.getRange(g.fila, 8).setValue("Enviado");

      pdfCursor += g.entradas;
      enviados++;
      enviadosMails.push(g.mail);
      Utilities.sleep(1200);
    } catch (err) {
      if (!sent) {
        // No se envió: liberar la fila y sus PDFs para el próximo intento
        try {
          hoja.getRange(g.fila, 8).setValue("Pendiente");
          hoja.getRange(g.fila, 9).setValue("");
        } catch (e2) {}
      } else {
        pdfCursor += g.entradas; // el mail salió: sus PDFs quedan consumidos
      }
      errores.push(g.nombre + " (" + g.mail + "): " + err.message);
    }
  }

  // restantes = pendientes NO intentados (corte por tiempo). Los que fallaron con
  // error NO cuentan: reintentarlos automáticamente podría loopear para siempre.
  const restantes = ganadores.length - idx;

  return {
    ok: true,
    enviados: enviados,
    enviadosMails: enviadosMails,
    errores: errores,
    restantes: restantes,
    mensaje: enviados + " mail" + (enviados !== 1 ? "s" : "") + " enviado" + (enviados !== 1 ? "s" : "") +
             (restantes ? " · quedan " + restantes + " pendientes" : "") +
             (errores.length ? " · " + errores.length + " con error" : "")
  };
}
```

- [ ] **Step 3 (Agente):** Verificación estática: releer la función y confirmar que (a) `restantes` es `ganadores.length - idx` y el `break` ocurre ANTES de procesar el ganador `idx`, (b) `pdfCursor` solo avanza cuando `sent === true` o tras éxito, (c) col I se escribe antes de `sendEmail` y se limpia si `!sent`.

- [ ] **Step 4 (Agente):** Commit.

```bash
git add sorteo_script.gs
git commit -m "feat(backend): envio por lotes (corte 4 min + restantes) y entradas por fila sin reusar PDFs"
```

---

## Task 3: Backend — `checkPDFs` calcula expected/found desde la planilla

**Files:**
- Modify: `sorteo_script.gs` — función `checkPDFs` (≈446-525)

**Interfaces:**
- Consumes: `_pdfsUsadosDelShow` (Task 2); col J de "Ganadores".
- Produces: respuesta con el MISMO shape de siempre (`status`, `expected`, `found`, `pendCount`, `entradasXGan`, `folderName`, `pdfNames`), pero `expected` = suma de entradas de las filas `Pendiente` del show (fallback `pendCount × entradasXGan` si la planilla no tiene filas), y `found`/`pdfNames` excluyen PDFs ya usados.

- [ ] **Step 1 (Agente):** En `checkPDFs`, justo después de las líneas `var epg = ...; var pc = ...; var expected = pc * epg; var buscar = ...;`, insertar la lectura de la planilla:

```js
  // Fuente de verdad: filas Pendiente de la pestaña Ganadores (suma de col J).
  // Si no hay filas para el show, cae al cálculo viejo pc*epg (p.ej. desincronización).
  var usados = {};
  try {
    var hojaG = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID).getSheetByName("Ganadores");
    if (hojaG) {
      var datosG = hojaG.getDataRange().getValues();
      usados = _pdfsUsadosDelShow(datosG, showId);
      var sumaEnt = 0, filasPend = 0;
      for (var gi = 1; gi < datosG.length; gi++) {
        if (String(datosG[gi][1]).trim() === String(showId).trim() && String(datosG[gi][7]).trim() === "Pendiente") {
          filasPend++;
          var entG = parseInt(datosG[gi][9]);
          sumaEnt += (isNaN(entG) || entG < 1) ? epg : entG;
        }
      }
      if (filasPend > 0) { expected = sumaEnt; pc = filasPend; }
    }
  } catch (eG) {}
```

- [ ] **Step 2 (Agente):** Al final de la función, donde arma la lista `pdfs` de la carpeta, filtrar usados antes del return final. Reemplazar:

```js
    pdfs.sort();

    return {
      ok: true,
      status: pdfs.length >= expected ? "ok" : "insufficient",
      expected: expected,
      found: pdfs.length,
      pendCount: pc,
      entradasXGan: epg,
      folderName: carpetaShow.getName(),
      pdfNames: pdfs.slice(0, 20)
    };
```

por:

```js
    pdfs.sort();
    var disponibles = pdfs.filter(function(nom) { return !usados[nom]; });

    return {
      ok: true,
      status: disponibles.length >= expected ? "ok" : "insufficient",
      expected: expected,
      found: disponibles.length,
      pendCount: pc,
      entradasXGan: epg,
      folderName: carpetaShow.getName(),
      pdfNames: disponibles.slice(0, 20)
    };
```

- [ ] **Step 3 (Agente):** Commit.

```bash
git add sorteo_script.gs
git commit -m "feat(backend): checkPDFs suma entradas pendientes de la planilla y descuenta PDFs usados"
```

---

## Task 4: Backend — `actualizarEntradas` + `getGanadoresShow` + casos del router

**Files:**
- Modify: `sorteo_script.gs` — funciones nuevas después de `guardarGanadores`; switch de `doPost` (≈98-118)

**Interfaces:**
- Consumes: col J (Task 1); celda B123 de "Tracking Ganadores" (patrón de `trackingGanadores`).
- Produces:
  - `actualizarEntradas(showId, mail, entradas, entradasPrev)` → `{ok:true, entradas:n, deltaTickets:d}` o `{ok:false, error}`. Solo toca filas `Pendiente`. Ajusta B123 por `n − old` donde `old` = col J si tiene valor (autoritativo), si no `entradasPrev` del request; sin `old` conocido no toca B123 y `deltaTickets` = 0.
  - `getGanadoresShow(showId)` → `{ok:true, ganadores:[{mail, nombre, estado, entradas|null}]}` (todas las filas del show).
  - Ambas son acciones **admin** (no se agregan a `PUBLIC_ACTIONS` → el router ya exige token PIN).

- [ ] **Step 1 (Agente):** Agregar después de `guardarGanadores`:

```js
// ============================================================
//  ACTUALIZAR ENTRADAS — edita la cantidad de un ganador Pendiente
//  Ajusta B123 (total historico de tickets) por el delta para que
//  el total del dashboard no quede desfasado tras la edicion.
// ============================================================
function actualizarEntradas(showId, mail, entradas, entradasPrev) {
  if (!showId || !mail) return { ok: false, error: "showId y mail requeridos" };
  var n = parseInt(entradas);
  if (isNaN(n) || n < 1 || n > 10) return { ok: false, error: "Cantidad inválida (1 a 10)" };

  var ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  var hoja = ss.getSheetByName("Ganadores");
  if (!hoja) return { ok: false, error: "No existe la pestaña Ganadores" };

  var datos = hoja.getDataRange().getValues();
  var m = String(mail).trim().toLowerCase();

  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() === String(showId).trim() &&
        String(datos[i][5]).trim().toLowerCase() === m &&
        String(datos[i][7]).trim() === "Pendiente") {

      var oldCell = parseInt(datos[i][9]);
      var old = (!isNaN(oldCell) && oldCell > 0) ? oldCell : (parseInt(entradasPrev) || null);

      hoja.getRange(i + 1, 10).setValue(n);

      var delta = 0;
      if (old !== null && old !== n) {
        delta = n - old;
        try {
          var tr = ss.getSheetByName("Tracking Ganadores");
          var valB = tr.getRange("B123").getValue();
          var base = (valB && !isNaN(Number(valB))) ? parseInt(valB) : 0;
          tr.getRange("B123").setValue(base + delta);
        } catch (e2) {
          Logger.log("actualizarEntradas: no se pudo ajustar B123: " + e2.message);
          delta = 0;
        }
      }
      return { ok: true, entradas: n, deltaTickets: delta };
    }
  }
  return { ok: false, error: "No hay fila Pendiente para ese ganador en este show (¿ya se envió?)" };
}

// ============================================================
//  GET GANADORES SHOW — filas de la pestaña Ganadores de un show
//  La app lo usa para hidratar cantidades al recargar (multi-compu).
// ============================================================
function getGanadoresShow(showId) {
  if (!showId) return { ok: false, error: "Show ID requerido" };
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  var hoja = ss.getSheetByName("Ganadores");
  if (!hoja) return { ok: true, ganadores: [] };

  var datos = hoja.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() !== String(showId).trim()) continue;
    var ent = parseInt(datos[i][9]);
    out.push({
      mail: String(datos[i][5]).trim(),
      nombre: String(datos[i][6]).trim(),
      estado: String(datos[i][7]).trim(),
      entradas: (isNaN(ent) || ent < 1) ? null : ent
    });
  }
  return { ok: true, ganadores: out };
}
```

- [ ] **Step 2 (Agente):** En el switch de `doPost`, agregar debajo del case `guardarGanadores`:

```js
      case "actualizarEntradas": return resp(actualizarEntradas(body.showId, body.mail, body.entradas, body.entradasPrev));
      case "getGanadoresShow":   return resp(getGanadoresShow(body.showId));
```

- [ ] **Step 3 (Agente):** Commit.

```bash
git add sorteo_script.gs
git commit -m "feat(backend): actualizarEntradas (con ajuste B123) y getGanadoresShow"
```

---

## Task 5: Backend — `trackingGanadores` suma por ganador + fix del parámetro dropeado en el router

**Files:**
- Modify: `sorteo_script.gs` — función `trackingGanadores` (≈750-751) y case del router (≈107)

**Interfaces:**
- Consumes: `body.ganadores` puede traer `entradas` por ganador (frontend nuevo, Task 7); `body.entradasXGan` como fallback.
- Produces: B123 += suma de `g.entradas` (fallback `entradasXGan` por ganador). `ticketsAdded` en la respuesta refleja la suma.

- [ ] **Step 1 (Agente):** En el router, reemplazar:

```js
      case "trackingGanadores": return resp(trackingGanadores(body.ganadores, body.showNombre, body.fecha));
```

por:

```js
      case "trackingGanadores": return resp(trackingGanadores(body.ganadores, body.showNombre, body.fecha, body.entradasXGan));
```

- [ ] **Step 2 (Agente):** En `trackingGanadores`, reemplazar:

```js
    var n = parseInt(entradasXGan) || 1;
    var ticketsAdded = ganadores.length * n;
```

por:

```js
    var n = parseInt(entradasXGan) || 1;
    var ticketsAdded = ganadores.reduce(function(acc, g) {
      var e = parseInt(g.entradas);
      return acc + ((isNaN(e) || e < 1) ? n : e);
    }, 0);
```

- [ ] **Step 3 (Agente):** Commit.

```bash
git add sorteo_script.gs
git commit -m "fix(backend): trackingGanadores suma entradas por ganador (router dropeaba entradasXGan)"
```

---

## Task 6 (Mateo): Deploy del backend + smoke test con el frontend viejo

**Files:** ninguno (operación en Google).

**Interfaces:**
- Consumes: `sorteo_script.gs` final de Tasks 1-5.
- Produces: deploy activo en la MISMA URL con el código nuevo; frontend actual (sin cambios) sigue funcionando.

- [ ] **Step 1 (Mateo):** Copiar todo `sorteo_script.gs` del repo y pegarlo en el editor de Apps Script (reemplaza el contenido). Guardar.

- [ ] **Step 2 (Mateo):** Publicar: Implementar → **Administrar implementaciones** → ✏️ en la implementación activa → Versión: **Nueva versión** → Implementar. Confirmar que la URL NO cambió y que sigue "Ejecutar como: Yo" + "Cualquier persona".

- [ ] **Step 3 (Mateo):** Smoke test de retro-compatibilidad en la app actual (GitHub Pages, sin cambios de frontend):
  1. Abrir la app, ingresar PIN, sincronizar un show.
  2. En la planilla, pestaña Ganadores: verificar que apareció el header **J1 = "Entradas"** después de la próxima confirmación de ganadores (o correr una confirmación de prueba).
  3. `checkPDFs` desde el botón "Enviar via Gmail" de un show ya enviado por completo: debe responder sin error (expected 0 o mensaje "ya enviaste").

Esperado: todo igual que antes; ninguna función rota.

---

## Task 7: Frontend — campo `entradas` al confirmar + payloads a Sheets

**Files:**
- Modify: `index.html` — funciones `confirmarGanadores` (≈3332), `guardarEnSheets` (≈3611), `trackingEnSheets` (≈3616)

**Interfaces:**
- Consumes: `ev.entradasXGan` (default del evento).
- Produces: cada ganador confirmado nace con `entradas: ev.entradasXGan || 1`; `guardarGanadores` y `trackingGanadores` reciben `entradas` por ganador.

- [ ] **Step 1 (Agente):** En `confirmarGanadores`, agregar `entradas` al objeto pusheado. Reemplazar:

```js
  lista.forEach(g=>S.ganadores.push({
    ...g, id:"g"+Date.now()+"_"+Math.floor(Math.random()*99999), evId, evNombre:ev.nombre,
    venue:ev.venue, fecha:ev.fecha, hora:ev.hora,
    estado:"pendiente", fechaGano:new Date().toISOString()
  }));
```

por:

```js
  lista.forEach(g=>S.ganadores.push({
    ...g, id:"g"+Date.now()+"_"+Math.floor(Math.random()*99999), evId, evNombre:ev.nombre,
    venue:ev.venue, fecha:ev.fecha, hora:ev.hora,
    entradas:ev.entradasXGan||1,
    estado:"pendiente", fechaGano:new Date().toISOString()
  }));
```

(El incremento local `S.ticketsBase += lista.length*n` más abajo queda como está: al confirmar todos tienen el default, la suma es equivalente.)

- [ ] **Step 2 (Agente):** En `guardarEnSheets`, reemplazar:

```js
  const res=await api({action:"guardarGanadores",ganadores:lista.map(g=>({mail:g.email,nombre:g.nombre})),showId:ev.id,showNombre:ev.nombre,fecha:fdL(ev.fecha),venue:ev.venue||"Movistar Arena"});
```

por:

```js
  const res=await api({action:"guardarGanadores",ganadores:lista.map(g=>({mail:g.email,nombre:g.nombre,entradas:ev.entradasXGan||1})),showId:ev.id,showNombre:ev.nombre,fecha:fdL(ev.fecha),venue:ev.venue||"Movistar Arena"});
```

- [ ] **Step 3 (Agente):** En `trackingEnSheets`, reemplazar:

```js
    ganadores:lista.map(g=>({nombre:g.nombre,mail:g.email})),
```

por:

```js
    ganadores:lista.map(g=>({nombre:g.nombre,mail:g.email,entradas:ev.entradasXGan||1})),
```

- [ ] **Step 4 (Agente):** Commit.

```bash
git add index.html
git commit -m "feat(entradas): ganadores nacen con entradas (default del evento) y viajan a Sheets"
```

- [ ] **Step 5 (Mateo):** Confirmación de prueba: sortear un show de prueba con 2 ganadores (evento con "entradas por ganador" = 2) → en la pestaña Ganadores, col J de las filas nuevas = 2; B123 subió +4.

---

## Task 8: Frontend — helper `entradasDe` + conteos por suma

**Files:**
- Modify: `index.html` — utils (junto a `const v=...`, ≈3910); `drawTickets` (dos definiciones: ≈1861 y ≈2228); `renderDash` (≈2670)

**Interfaces:**
- Consumes: `g.entradas` (puede faltar en ganadores viejos/reimportados), `ev.entradasXGan`.
- Produces: `entradasDe(g)` global → entero ≥ 1 (prioridad: `g.entradas` → `ev.entradasXGan` → 1). Todos los conteos de tickets lo usan.

- [ ] **Step 1 (Agente):** Agregar junto a los utils (antes de `const v=...`):

```js
function entradasDe(g){
  const n=parseInt(g.entradas);
  if(!isNaN(n)&&n>0)return n;
  const ev=S.eventos.find(e=>e.id===g.evId);
  return ev?.entradasXGan||1;
}
```

- [ ] **Step 2 (Agente):** En la **primera** `drawTickets` (≈1861), reemplazar:

```js
  gans.forEach(g=>{
    const ev=S.eventos.find(e=>e.id===g.evId);
    const n=ev?.entradasXGan||1;
    if(g.estado==="enviado")totalEnv+=n;
    else totalPend+=n;
  });
```

por:

```js
  gans.forEach(g=>{
    const n=entradasDe(g);
    if(g.estado==="enviado")totalEnv+=n;
    else totalPend+=n;
  });
```

- [ ] **Step 3 (Agente):** En la **segunda** `drawTickets` (≈2228, es la que gana por redefinición), reemplazar:

```js
  S.ganadores.forEach(function(g){
    if(g._historico)return; // Skip historical ganadores - already counted in base
    var ev=S.eventos.find(function(e){return e.id===g.evId;});
    var n=ev&&ev.entradasXGan?ev.entradasXGan:1;
    if(g.estado==="enviado")totalEnv+=n;
    else totalPend+=n;
  });
```

por:

```js
  S.ganadores.forEach(function(g){
    if(g._historico)return; // Skip historical ganadores - already counted in base
    var n=entradasDe(g);
    if(g.estado==="enviado")totalEnv+=n;
    else totalPend+=n;
  });
```

- [ ] **Step 4 (Agente):** En `renderDash`, reemplazar:

```js
  const totalTick=S.ganadores.reduce((acc,g)=>{const ev=S.eventos.find(e=>e.id===g.evId);return acc+(ev?.entradasXGan||1);},0);
  const tickEnv=S.ganadores.filter(g=>g.estado==="enviado").reduce((acc,g)=>{const ev=S.eventos.find(e=>e.id===g.evId);return acc+(ev?.entradasXGan||1);},0);
```

por:

```js
  const totalTick=S.ganadores.reduce((acc,g)=>acc+entradasDe(g),0);
  const tickEnv=S.ganadores.filter(g=>g.estado==="enviado").reduce((acc,g)=>acc+entradasDe(g),0);
```

- [ ] **Step 5 (Agente):** Commit.

```bash
git add index.html
git commit -m "feat(entradas): conteos de tickets suman por ganador via entradasDe()"
```

- [ ] **Step 6 (Mateo):** En la app recargada, F12 → Console:

```js
typeof entradasDe // "function"
entradasDe({entradas:4}) // 4
entradasDe({evId:"NO_EXISTE"}) // 1
```

y verificar que el donut de tickets y los KPI del dashboard muestran los mismos números que antes (nada cambió aún porque todos tienen el default).

---

## Task 9: Frontend — stepper − / + en ganadores confirmados + cantidad en las cards de mail

**Files:**
- Modify: `index.html` — `renderGanConf` (≈3648-3656), `renderMails` (≈3737-3763); funciones nuevas `stepEntradas` / `_pushEntradas` junto a `renderGanConf`

**Interfaces:**
- Consumes: `entradasDe(g)` (Task 8); endpoint `actualizarEntradas` (Task 4) vía `api({action:"actualizarEntradas", showId, mail, entradas, entradasPrev})`.
- Produces: `stepEntradas(gid, delta)` global (onclick). Ajuste local inmediato + push al backend con debounce 600 ms; si falla, revierte y avisa. `res.deltaTickets` se aplica a `S.ticketsBase` solo si el ganador es `_historico` (ya contado en la base).

- [ ] **Step 1 (Agente):** En `renderGanConf`, reemplazar el template de fila:

```js
  const rows=gans.map((g,i)=>{
    const canDel=g.estado!=="enviado";
    return `<div style="display:flex;align-items:center;gap:9px;padding:8px 11px;border-bottom:1px solid var(--line)">
      <div style="font-family:var(--fd);font-size:16px;color:var(--ice);opacity:.35;width:22px;text-align:center;line-height:1">${i+1}</div>
      <div style="flex:1;min-width:0"><div style="font-size:12px;font-weight:500">${g.nombre}</div><div style="font-size:10px;color:var(--gray3)">${g.email}${g.dni?" · DNI "+g.dni:""}</div></div>
      <span class="bic ${g.estado==="enviado"?"bgi":"bai"}">${g.estado==="enviado"?"✓ Enviado":"Pendiente"}</span>
      ${canDel?`<button class="btn br bsm" style="margin-left:4px" onclick="eliminarGanConfirmado('${g.id}')" title="Eliminar de la lista">✕</button>`:''}
    </div>`;
  }).join("");
```

por:

```js
  const rows=gans.map((g,i)=>{
    const canDel=g.estado!=="enviado";
    const ent=entradasDe(g);
    const entCtl=g.estado==="enviado"
      ?`<span style="font-size:11px;color:var(--gray3);flex-shrink:0;min-width:44px;text-align:center">🎟 ${ent}</span>`
      :`<div style="display:flex;align-items:center;gap:3px;flex-shrink:0">
          <button class="btn bo bsm" style="padding:2px 8px" data-gid="${g.id}" onclick="stepEntradas(this.dataset.gid,-1)" title="Menos entradas">−</button>
          <span style="font-size:11px;color:var(--ice2);min-width:34px;text-align:center">🎟 ${ent}</span>
          <button class="btn bo bsm" style="padding:2px 8px" data-gid="${g.id}" onclick="stepEntradas(this.dataset.gid,1)" title="Más entradas">+</button>
        </div>`;
    return `<div style="display:flex;align-items:center;gap:9px;padding:8px 11px;border-bottom:1px solid var(--line)">
      <div style="font-family:var(--fd);font-size:16px;color:var(--ice);opacity:.35;width:22px;text-align:center;line-height:1">${i+1}</div>
      <div style="flex:1;min-width:0"><div style="font-size:12px;font-weight:500">${g.nombre}</div><div style="font-size:10px;color:var(--gray3)">${g.email}${g.dni?" · DNI "+g.dni:""}</div></div>
      ${entCtl}
      <span class="bic ${g.estado==="enviado"?"bgi":"bai"}">${g.estado==="enviado"?"✓ Enviado":"Pendiente"}</span>
      ${canDel?`<button class="btn br bsm" style="margin-left:4px" onclick="eliminarGanConfirmado('${g.id}')" title="Eliminar de la lista">✕</button>`:''}
    </div>`;
  }).join("");
```

- [ ] **Step 2 (Agente):** Agregar después de `renderGanConf`:

```js
const _entTimers={};
function stepEntradas(gid,delta){
  const g=S.ganadores.find(x=>String(x.id)===String(gid));
  if(!g||g.estado==="enviado")return;
  const actual=entradasDe(g);
  const nuevo=Math.min(10,Math.max(1,actual+parseInt(delta)));
  if(nuevo===actual)return;
  if(g._entPrev===undefined)g._entPrev=actual; // base del delta hasta que el backend confirme
  g.entradas=nuevo;
  save();renderGanConf();renderDash();
  clearTimeout(_entTimers[gid]);
  _entTimers[gid]=setTimeout(()=>_pushEntradas(g),600);
}
async function _pushEntradas(g){
  const prev=g._entPrev;delete g._entPrev;
  const ev=S.eventos.find(e=>e.id===g.evId);
  try{
    const res=await api({action:"actualizarEntradas",showId:String(ev?.id||g.evId),mail:g.email,entradas:g.entradas,entradasPrev:prev});
    if(!res.ok)throw new Error(res.error||"sin respuesta");
    if(typeof res.deltaTickets==="number"&&res.deltaTickets!==0&&g._historico){
      S.ticketsBase=(S.ticketsBase||0)+res.deltaTickets;
      save();renderDash();
    }
    toast("🎟 "+g.nombre.split(" ")[0]+": "+g.entradas+" entrada(s)");
  }catch(err){
    if(prev!==undefined){g.entradas=prev;save();renderGanConf();renderDash();}
    toast("No se pudo actualizar entradas: "+err.message,"warn");
  }
}
```

- [ ] **Step 3 (Agente):** En `renderMails`, dentro del bloque del destinatario (después de la línea del DNI), agregar la cantidad como texto. Reemplazar:

```js
              ${g.dni?`<div style="font-size:10px;color:var(--gray3);margin-top:1px">DNI: ${g.dni}</div>`:""}
```

por:

```js
              ${g.dni?`<div style="font-size:10px;color:var(--gray3);margin-top:1px">DNI: ${g.dni}</div>`:""}
              <div style="font-size:10px;color:var(--gray3);margin-top:1px">🎟 ${entradasDe(g)} entrada${entradasDe(g)!==1?"s":""}</div>
```

- [ ] **Step 4 (Agente):** Commit.

```bash
git add index.html
git commit -m "feat(entradas): stepper por ganador con push a Sheets (actualizarEntradas)"
```

- [ ] **Step 5 (Mateo):** En el show de prueba de la Task 7: subir un ganador de 2 → 4 con el stepper. Esperado: la fila cambia a 🎟 4, toast de confirmación, col J = 4 en la planilla, B123 subió +2, y el donut de tickets refleja el cambio. Después bajarlo a 1 y verificar el camino inverso. Por último, probar el stepper en un ganador de un show viejo ya enviado (reimportado del Tracking): debe dar toast de error "No hay fila Pendiente..." y volver al valor anterior.

---

## Task 10: Frontend — hidratación de cantidades desde la planilla (multi-compu)

**Files:**
- Modify: `index.html` — función nueva `hydrateEntradas` junto a `renderGanConf`; llamadas al inicio de `renderGanConf` (≈3629) y `renderStep4` (≈3694)

**Interfaces:**
- Consumes: `getGanadoresShow` (Task 4); `S.ganadores` (los reimportados del Tracking tienen mail sintetizado → se matchea por mail O por nombre en mayúsculas).
- Produces: `hydrateEntradas(evId)` async, cacheada por `Set` por sesión; tras mergear re-renderiza.

- [ ] **Step 1 (Agente):** Agregar antes de `renderGanConf`:

```js
const _entHydrated=new Set();
async function hydrateEntradas(evId){
  if(!evId||_entHydrated.has(evId))return;
  _entHydrated.add(evId);
  const ev=S.eventos.find(e=>e.id===evId);
  try{
    const res=await api({action:"getGanadoresShow",showId:String(ev?.id||evId)});
    if(!res.ok||!Array.isArray(res.ganadores))return;
    let changed=false;
    res.ganadores.forEach(r=>{
      if(!r.entradas)return;
      const rMail=String(r.mail||"").trim().toLowerCase();
      const rNom=String(r.nombre||"").trim().toUpperCase();
      const g=S.ganadores.find(x=>x.evId===evId&&(
        String(x.email||"").trim().toLowerCase()===rMail||
        String(x.nombre||"").trim().toUpperCase()===rNom
      ));
      if(g&&g.entradas!==r.entradas){g.entradas=r.entradas;changed=true;}
    });
    if(changed){save();renderGanConf();if(document.getElementById("mail-list"))renderMails();renderDash();}
  }catch(e){
    _entHydrated.delete(evId); // permitir reintento en el próximo render
  }
}
```

- [ ] **Step 2 (Agente):** Primera línea del cuerpo de `renderGanConf`: agregar `hydrateEntradas(S.evActivo);` (fire-and-forget; el cache evita spam). Ídem en `renderStep4`, después de `const evId=S.evActivo;` agregar `hydrateEntradas(evId);`.

- [ ] **Step 3 (Agente):** Commit.

```bash
git add index.html
git commit -m "feat(entradas): hidratar cantidades desde la pestana Ganadores al abrir el show (multi-compu)"
```

- [ ] **Step 4 (Mateo):** Con el show de prueba: dejar un ganador en 🎟 3, recargar la app (F5), abrir el paso de ganadores del show. Esperado: tras ~1-2 s la fila muestra 🎟 3 (vino de la planilla, no del default). Ideal: repetir desde la compu del otro usuario — debe ver 3 también.

---

## Task 11: Frontend — envío con auto-reanudación y progreso

**Files:**
- Modify: `index.html` — `_enviarMailsCore` (≈3839-3856); texto "insufficient" de `mostrarPreflightPDFs` (≈3878-3881)

**Interfaces:**
- Consumes: respuesta de `enviarMails` con `restantes` (Task 2); `pend` (array local de pendientes del show).
- Produces: loop `do/while` que re-llama mientras `restantes > 0` (tope 30 vueltas), progreso en el botón, errores acumulados entre vueltas. La alerta de desincronización solo corre en la vuelta 1.

- [ ] **Step 1 (Agente):** Reemplazar `_enviarMailsCore` completa por:

```js
async function _enviarMailsCore(showKey,showName,ev,pend){
  const btn=document.getElementById("btn-gmail");
  btn.disabled=true;
  const total=pend.length;
  let acumEnviados=0,vuelta=0,res=null,desync=false;
  const erroresAcum=[];
  try{
    do{
      vuelta++;
      btn.textContent="⏳ Enviando… "+acumEnviados+"/"+total;
      res=await api({action:"enviarMails",showId:String(showKey),entradasXGan:ev?.entradasXGan||1});
      if(!res.ok){toast("Error: "+(res.error||"sin respuesta"),"warn");break;}
      // Desincronización (solo vuelta 1): esperábamos pendientes pero la planilla no devolvió ninguno
      if(vuelta===1&&res.enviados===0&&!(res.errores&&res.errores.length)&&!res.restantes&&pend.length>0){
        desync=true;
        alert("Esperaba "+pend.length+" ganador(es) pendiente(s) para este show, pero la planilla no devolvió ninguno.\n\nPosible desincronización — revisá/reintentá la sincronización antes de enviar.");
        break;
      }
      // Marcar localmente solo los confirmados por el backend (si no trae lista, marcar todos)
      const okMails=Array.isArray(res.enviadosMails)?new Set(res.enviadosMails.map(m=>String(m).trim().toLowerCase())):null;
      pend.forEach(g=>{if(g.estado!=="enviado"&&(!okMails||okMails.has(String(g.email||g.mail||"").trim().toLowerCase())))g.estado="enviado";});
      acumEnviados+=res.enviados||0;
      if(res.errores?.length)erroresAcum.push(...res.errores);
      save();renderMails();renderDash();
      btn.textContent="⏳ Enviando… "+acumEnviados+"/"+total;
    }while(res.ok&&res.restantes>0&&vuelta<30);
  }finally{
    btn.disabled=false;btn.textContent="Enviar via Gmail";
  }
  if(res&&res.ok&&!desync){
    toast(acumEnviados+" mail"+(acumEnviados!==1?"s":"")+" enviado"+(acumEnviados!==1?"s":"")+(erroresAcum.length?" · "+erroresAcum.length+" con error":""));
  }
  if(erroresAcum.length)setTimeout(()=>alert("Errores:\n"+erroresAcum.join("\n")),500);
}
```

Notas de diseño que el implementador debe respetar:
- `pend` son los pendientes locales al momento de apretar el botón; se marcan `enviado` solo los mails confirmados por el backend en CADA vuelta.
- Si una vuelta falla (`!res.ok` o excepción de red), lo ya marcado queda marcado; apretar "Enviar via Gmail" de nuevo **reanuda** (el backend solo toma filas `Pendiente`). No hay botón aparte de "Reanudar".
- El caso "0 pendientes y sin restantes en vuelta 1" mantiene la alerta de desincronización existente.

- [ ] **Step 2 (Agente):** En `mostrarPreflightPDFs`, reemplazar el detalle del caso `insufficient`:

```js
    detalle="Necesitás <strong>"+check.expected+"</strong> PDFs ("+check.pendCount+" ganador"+(check.pendCount!==1?"es":"")+" × "+check.entradasXGan+" entrada"+(check.entradasXGan!==1?"s":"")+") pero hay solo <strong>"+check.found+"</strong> en la carpeta <em>'"+esc(check.folderName||"")+"'</em>.<br>Subí los PDFs que faltan y reintentá.";
```

por:

```js
    detalle="Necesitás <strong>"+check.expected+"</strong> PDF(s) para los "+check.pendCount+" ganador"+(check.pendCount!==1?"es":"")+" pendiente"+(check.pendCount!==1?"s":"")+" pero hay solo <strong>"+check.found+"</strong> disponibles en la carpeta <em>'"+esc(check.folderName||"")+"'</em> (los ya enviados a otros ganadores no cuentan).<br>Subí los PDFs que faltan y reintentá.";
```

(Con cantidades por ganador la cuenta "N × M" ya no aplica; `expected` viene sumado desde la planilla.)

- [ ] **Step 3 (Agente):** Commit.

```bash
git add index.html
git commit -m "feat(envio): loop de auto-reanudacion con progreso usando restantes del backend"
```

---

## Task 12 (Mateo): Verificación integral + checklist multi-usuario

Corresponde a la sección "Pruebas" de la spec. Usar un show de prueba con mails propios (los tuyos y el del otro usuario), nunca empleados reales.

- [ ] **Prueba 1 — cantidades mixtas:** show de prueba con 3 ganadores, ajustar a 2 / 2 / 4 con el stepper. Carpeta Drive con 7 PDFs → "Enviar via Gmail" debe bloquear con "Necesitás 8 PDF(s)… hay solo 7". Subir el 8º → envía; verificar en los mails recibidos 2, 2 y 4 adjuntos SIN repetidos, y col I de la planilla con los nombres correctos.

- [ ] **Prueba 2 — reanudación sin duplicados:** en el editor de Apps Script, bajar temporalmente `BUDGET_MS` a `10000` (10 s) y publicar nueva versión. Show de prueba con 4+ ganadores pendientes → "Enviar": la app debe encadenar 2+ vueltas mostrando "X/N enviados…" y terminar completo. Verificar que ningún mail llegó dos veces y ningún PDF se repitió entre destinatarios. **Restaurar `BUDGET_MS = 240000` y publicar nueva versión.**

- [ ] **Prueba 3 — persistencia multi-compu:** ajustar un pendiente a 🎟 3 en la compu A → en la compu B (o ventana incógnito con el PIN), abrir el show: debe mostrar 🎟 3. Enviar desde B → el mail trae 3 PDFs.

- [ ] **Prueba 4 — legacy:** una fila vieja de Ganadores sin col J (pre-cambio) pendiente → enviar: usa el `entradasXGan` del evento como siempre.

- [ ] **Prueba 5 — checklist multi-usuario (con el otro usuario):**
  1. Abre la app en su compu e ingresa el PIN.
  2. Ajusta una cantidad → se ve en la planilla y en la compu de Mateo al recargar.
  3. Dispara un envío de prueba completo → los mails salen desde `rrhh@buenosairesarena.com`.
  4. Simulacro de reanudación cruzada: envío cortado en una compu (Ctrl+W a mitad del progreso) se retoma desde la otra apretando "Enviar via Gmail", sin duplicados.

- [ ] **Cierre:** si todo pasa, actualizar `CONTEXTO_PROYECTO.md` (nuevos endpoints `actualizarEntradas` / `getGanadoresShow`, campo `entradas` del ganador, `restantes` en `enviarMails`) y commitear:

```bash
git add CONTEXTO_PROYECTO.md
git commit -m "docs: contexto actualizado — entradas por ganador y envio por lotes"
```
