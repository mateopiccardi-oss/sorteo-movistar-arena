# Autenticación del backend + enviarMails robusto — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Proteger el Apps Script público con un token derivado del PIN y hacer que `enviarMails` no pueda duplicar mails ni enviarlos sin entrada adjunta.

**Architecture:** El backend valida un `token` (SHA-256 hex del PIN, almacenado en Script Properties) en todas las acciones salvo las 4 públicas del formulario. El frontend valida el PIN contra el servidor y la función `api()` inyecta el token en cada request. `enviarMails` se envuelve en `LockService`, pre-valida la cantidad de PDFs y marca cada fila "Enviando" antes de mandar el mail.

**Tech Stack:** Google Apps Script (V8), HTML/JS vanilla. Sin framework de tests — Apps Script solo corre en Google; cada task incluye verificación estática (`node --check`) y el spec define el plan de prueba manual post-deploy.

**Spec:** `docs/superpowers/specs/2026-06-11-auth-backend-envio-robusto-design.md`

---

### Task 1: Backend — auth por token

**Files:**
- Modify: `sorteo_script.gs:59-90` (router) y agregar bloque AUTH después de CONFIG (línea ~57)

- [ ] **Step 1: Agregar bloque AUTH después del cierre de CONFIG (línea 57)**

```javascript
// ============================================================
//  AUTH — acciones públicas vs. admin
//  El token es el SHA-256 hex del PIN, guardado en Script Properties.
//  Setup: ejecutar configurarPin() una vez desde el editor.
// ============================================================
const PUBLIC_ACTIONS = ["validarMail", "inscribir", "checkShowActivo", "validarPin"];

function _authOk(body) {
  const stored = PropertiesService.getScriptProperties().getProperty("ADMIN_HASH");
  if (!stored) return false; // sin PIN configurado → acciones admin bloqueadas
  return !!(body && body.token === stored);
}

function _sha256hex(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(function(b) { return ((b + 256) % 256).toString(16).padStart(2, "0"); }).join("");
}

// SETUP (una sola vez): ① escribí el PIN nuevo acá, ② Ejecutar → configurarPin
// desde el editor de Apps Script, ③ volvé a dejar el placeholder y guardá.
function configurarPin() {
  const PIN = "PONE_EL_PIN_ACA";
  if (PIN === "PONE_EL_PIN_ACA") throw new Error("Editá la constante PIN con el PIN nuevo antes de ejecutar.");
  PropertiesService.getScriptProperties().setProperty("ADMIN_HASH", _sha256hex(PIN));
  Logger.log("PIN configurado correctamente.");
}
```

- [ ] **Step 2: Proteger el router en `doPost` (líneas 62-67)**

Reemplazar:

```javascript
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    switch (action) {
      case "validarMail":       return resp(validarMail(body.mail));
```

por:

```javascript
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (PUBLIC_ACTIONS.indexOf(action) === -1 && !_authOk(body)) {
      return resp({ ok: false, error: "No autorizado" });
    }

    switch (action) {
      case "validarPin":        return resp({ ok: true, valido: _authOk(body) });
      case "validarMail":       return resp(validarMail(body.mail));
```

- [ ] **Step 3: Verificación estática**

Run: `Copy-Item sorteo_script.gs $env:TEMP\sorteo_check.js; node --check $env:TEMP\sorteo_check.js`
Expected: sin output (sintaxis OK)

- [ ] **Step 4: Commit**

```bash
git add sorteo_script.gs
git commit -m "feat(backend): exigir token en acciones admin del router"
```

---

### Task 2: Backend — enviarMails con lock, pre-check de PDFs e idempotencia

**Files:**
- Modify: `sorteo_script.gs:270-356` (función `enviarMails` completa)

- [ ] **Step 1: Reemplazar la función `enviarMails` entera (líneas 270-356) por:**

```javascript
function enviarMails(showId, entradasXGan) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { ok: false, error: "Otro envío está en curso — esperá unos segundos y reintentá." };
  }
  try {
    return _enviarMailsLocked(showId, entradasXGan);
  } finally {
    lock.releaseLock();
  }
}

function _enviarMailsLocked(showId, entradasXGan) {
  if (!showId) return { ok: false, error: "Show ID requerido" };
  entradasXGan = parseInt(entradasXGan) || 1;

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_SORTEO_ID);
  const hoja = ss.getSheetByName("Ganadores");
  if (!hoja) return { ok: false, error: "No hay ganadores registrados" };

  const datos = hoja.getDataRange().getValues();
  const ganadores = [];

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() === String(showId).trim() && String(datos[i][7]).trim() === "Pendiente") {
      ganadores.push({
        mail: String(datos[i][5]).trim(),
        nombre: String(datos[i][6]).trim(),
        showNombre: String(datos[i][2]).trim(),
        fecha: String(datos[i][3]).trim(),
        venue: String(datos[i][4]).trim(),
        fila: i + 1,
      });
    }
  }

  if (!ganadores.length) return { ok: true, enviados: 0, enviadosMails: [], mensaje: "No hay ganadores pendientes" };

  const pdfs = buscarPDFs(showId, ganadores[0].showNombre);

  // Regla estricta: si los PDFs no alcanzan, no se envía NINGÚN mail
  const necesarios = ganadores.length * entradasXGan;
  if (pdfs.length < necesarios) {
    return {
      ok: false,
      error: "Hay " + pdfs.length + " PDF(s) en Drive y se necesitan " + necesarios +
             " (" + ganadores.length + " ganador(es) × " + entradasXGan + " entrada(s)) — no se envió ningún mail."
    };
  }

  const errores = [];
  const enviadosMails = [];
  let enviados = 0;

  ganadores.forEach((g, i) => {
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

      const pdfStart = i * entradasXGan;
      const pdfSlice = pdfs.slice(pdfStart, pdfStart + entradasXGan);
      opts.attachments = pdfSlice.map(function(p) { return p.getAs(MimeType.PDF); });

      // Enviar como HTML para soportar UTF-8 y emojis correctamente
      opts.htmlBody = cuerpo.replace(/\n/g, "<br>");

      // Marcar ANTES de enviar: si la ejecución se corta acá, un reintento
      // solo retoma filas "Pendiente" y nunca duplica este mail
      hoja.getRange(g.fila, 8).setValue("Enviando");
      GmailApp.sendEmail(g.mail, asunto, cuerpo, opts);
      sent = true;
      hoja.getRange(g.fila, 8).setValue("Enviado");
      hoja.getRange(g.fila, 9).setValue(pdfSlice.map(function(p) { return p.getName(); }).join(", "));

      enviados++;
      enviadosMails.push(g.mail);
      Utilities.sleep(1200);
    } catch (err) {
      if (!sent) {
        try { hoja.getRange(g.fila, 8).setValue("Pendiente"); } catch (e2) {}
      }
      errores.push(g.nombre + " (" + g.mail + "): " + err.message);
    }
  });

  return {
    ok: true,
    enviados: enviados,
    enviadosMails: enviadosMails,
    errores: errores,
    mensaje: enviados + " mail" + (enviados !== 1 ? "s" : "") + " enviado" + (enviados !== 1 ? "s" : "") +
             (errores.length ? " · " + errores.length + " con error" : "")
  };
}
```

Notas: el comentario-cabecera existente de la sección (líneas 258-269) se conserva. Desaparece el caso "Sin PDF" en la columna 9 porque el pre-check garantiza PDFs suficientes.

- [ ] **Step 2: Verificación estática**

Run: `Copy-Item sorteo_script.gs $env:TEMP\sorteo_check.js; node --check $env:TEMP\sorteo_check.js`
Expected: sin output (sintaxis OK)

- [ ] **Step 3: Commit**

```bash
git add sorteo_script.gs
git commit -m "feat(backend): enviarMails con lock, pre-check de PDFs y marcado Enviando->Enviado"
```

---

### Task 3: Frontend — token en api() y PIN validado contra el servidor

**Files:**
- Modify: `index.html:1232-1244` (función `api`)
- Modify: `index.html:3896-3914` (bloque AUTH)
- Modify: `index.html:4142` (gate de inicio)

- [ ] **Step 1: Inyectar token en `api()` (línea 1232)**

Reemplazar:

```javascript
async function api(data){
  try{
    const r=await fetch(SCRIPT_URL,{method:"POST",body:JSON.stringify(data)});
```

por:

```javascript
async function api(data){
  try{
    const tok=sessionStorage.getItem("ma_tok");
    const payload=tok?{...data,token:tok}:data;
    const r=await fetch(SCRIPT_URL,{method:"POST",body:JSON.stringify(payload)});
```

- [ ] **Step 2: Reemplazar el bloque AUTH (líneas 3896-3914)**

Reemplazar desde `// ── AUTH ──` hasta el cierre de `checkPin()` por:

```javascript
// ── AUTH ──────────────────────────────────────────────────
// El PIN se valida contra el servidor (Script Properties) — no hay hash en el código fuente.
async function _h(s){const b=await crypto.subtle.digest("SHA-256",new TextEncoder().encode(s));return Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join("");}
async function checkPin(){
  const inp=document.getElementById("pin-input");
  const errEl=document.getElementById("pin-error");
  const h=await _h(inp.value);
  inp.disabled=true;
  errEl.textContent="Verificando…";
  const res=await api({action:"validarPin",token:h});
  inp.disabled=false;
  if(res.ok&&res.valido){
    sessionStorage.setItem("ma_auth","1");
    sessionStorage.setItem("ma_tok",h);
    const gate=document.getElementById("pin-gate");
    gate.classList.add("out");
    setTimeout(()=>gate.remove(),320);
    startApp();
  }else{
    inp.classList.add("err");
    errEl.textContent=res.ok?"PIN incorrecto":"No se pudo verificar — revisá tu conexión";
    inp.value="";
    inp.focus();
    setTimeout(()=>{inp.classList.remove("err");errEl.textContent="";},1800);
  }
}
```

(Se elimina la constante `_PH`.)

- [ ] **Step 3: Gate de inicio exige también el token (línea 4142)**

Reemplazar:

```javascript
if(sessionStorage.getItem("ma_auth")==="1"){
```

por:

```javascript
if(sessionStorage.getItem("ma_auth")==="1"&&sessionStorage.getItem("ma_tok")){
```

- [ ] **Step 4: Verificación estática del JS embebido**

Run (PowerShell): extraer el contenido del último `<script>` de index.html a `$env:TEMP\idx_check.js` y correr `node --check`.
Expected: sin output (sintaxis OK)

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat(admin): validar PIN contra el servidor e inyectar token en api()"
```

---

### Task 4: Frontend — envío alineado con el backend estricto

**Files:**
- Modify: `index.html:3819-3828` (`_enviarMailsCore`)
- Modify: `index.html:3829-3878` (`mostrarPreflightPDFs`)

- [ ] **Step 1: Marcar localmente solo los mails confirmados (líneas 3824-3826)**

Reemplazar:

```javascript
  if(!res.ok){toast("Error: "+(res.error||"sin respuesta"),"warn");return;}
  pend.forEach(g=>{g.estado="enviado";});
  save();toast(res.mensaje||pend.length+" mails enviados");renderMails();renderDash();
```

por:

```javascript
  if(!res.ok){toast("Error: "+(res.error||"sin respuesta"),"warn");return;}
  const okMails=Array.isArray(res.enviadosMails)?new Set(res.enviadosMails.map(m=>String(m).trim().toLowerCase())):null;
  pend.forEach(g=>{if(!okMails||okMails.has(String(g.email||g.mail||"").trim().toLowerCase()))g.estado="enviado";});
  save();toast(res.mensaje||pend.length+" mails enviados");renderMails();renderDash();
```

(`okMails===null` mantiene el comportamiento actual si el backend todavía no está redeployado.)

- [ ] **Step 2: Quitar "Enviar igual" del preflight (líneas 3832, 3853, 3855-3857 y 3872-3877)**

a) Línea 3832, reemplazar:

```javascript
  const blocker=(check.status==="no-root"||check.status==="no-entradas");
```

por:

```javascript
  const blocker=true; // el backend rechaza envíos con PDFs insuficientes — no hay "enviar igual"
```

b) Línea 3853, reemplazar el final del texto de `insufficient`:

```javascript
    detalle="Necesitás <strong>"+check.expected+"</strong> PDFs ("+check.pendCount+" ganador"+(check.pendCount!==1?"es":"")+" × "+check.entradasXGan+" entrada"+(check.entradasXGan!==1?"s":"")+") pero hay solo <strong>"+check.found+"</strong> en la carpeta <em>'"+esc(check.folderName||"")+"'</em>.<br>Podés subir más PDFs y reintentar, o enviar igual (los últimos ganadores se quedarán sin adjunto).";
```

por:

```javascript
    detalle="Necesitás <strong>"+check.expected+"</strong> PDFs ("+check.pendCount+" ganador"+(check.pendCount!==1?"es":"")+" × "+check.entradasXGan+" entrada"+(check.entradasXGan!==1?"s":"")+") pero hay solo <strong>"+check.found+"</strong> en la carpeta <em>'"+esc(check.folderName||"")+"'</em>.<br>Subí los PDFs que faltan y reintentá.";
    color="var(--red)";
```

c) Líneas 3855-3857, reemplazar:

```javascript
  const acciones=blocker
    ? '<button class="btn bo" onclick="document.getElementById(\'preflight-modal\').remove()">Cerrar</button>'
    : '<button class="btn bo" onclick="document.getElementById(\'preflight-modal\').remove()">Cancelar</button><button class="btn br" id="btn-pf-force">Enviar igual</button>';
```

por:

```javascript
  const acciones='<button class="btn bo" onclick="document.getElementById(\'preflight-modal\').remove()">Cerrar</button>';
```

d) Líneas 3872-3877, eliminar el bloque:

```javascript
  if(!blocker){
    document.getElementById("btn-pf-force").addEventListener("click",async()=>{
      modal.remove();
      await _enviarMailsCore(nombreCarpeta,ev,pend);
    });
  }
```

y como `blocker` queda sin usos, eliminar también la línea agregada en (a) — `const blocker=true;` — quedando el modal siempre con "Cerrar".

- [ ] **Step 3: Verificación estática del JS embebido**

Run (PowerShell): extraer el `<script>` principal de index.html a `$env:TEMP\idx_check.js` y correr `node --check`.
Expected: sin output (sintaxis OK)

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat(admin): quitar 'Enviar igual' y marcar enviados segun respuesta del backend"
```

---

### Task 5: Deploy y prueba manual (lo hace Mateo)

- [ ] Pegar `sorteo_script.gs` actualizado en el editor de Apps Script.
- [ ] Ejecutar `configurarPin()` con el PIN nuevo (8+ dígitos) y restaurar el placeholder.
- [ ] Implementar → Administrar implementaciones → ✏️ → **Versión nueva** (NO crear implementación nueva; modo queda "Yo + Cualquier persona").
- [ ] Push a GitHub (Pages se publica solo) y verificar el deploy de Netlify.
- [ ] Correr el plan de prueba manual del spec (PIN, "No autorizado" sin token, formulario sigue andando, envío con PDFs de menos, reintento sin duplicados).
- [ ] Avisar el PIN nuevo a la otra usuaria.
