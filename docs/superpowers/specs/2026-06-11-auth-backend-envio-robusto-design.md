# Diseño: Autenticación del backend + envío de mails robusto

**Fecha:** 2026-06-11
**Estado:** Aprobado

## Contexto

El Apps Script está deployado como "Ejecutar como: Yo + Cualquier persona" (no se puede cambiar — rompe CORS). La URL es visible en el código fuente del formulario público y de la app admin, y el router `doPost` no valida quién llama: cualquiera con la URL puede enviar mails masivos, escribir el Tracking, borrar shows o descargar datos de empleados. Además, `enviarMails` marca "Enviado" después de mandar cada mail, así que un reintento tras un corte puede duplicar mails, y si faltan PDFs en Drive el mail sale igual sin adjunto.

## Decisiones tomadas

- **PIN numérico de 8+ dígitos** (elegido por el usuario), compartido entre las 2 usuarias de RRHH, igual que hoy. El PIN nuevo lo configura Mateo directo en el editor de Apps Script; nunca queda en el repo.
- **Si faltan PDFs, no se envía ningún mail** (estricto). Se elimina la opción "Enviar igual" del preflight del frontend porque el servidor ahora la rechazaría.

## Mejora 1 — Autenticación por token

### Backend (`sorteo_script.gs`)

- **Acciones públicas** (sin token): `validarMail`, `inscribir`, `checkShowActivo` (las 3 que usa el formulario de empleados) + `validarPin` (nueva).
- **Todo lo demás exige token**: en `doPost`, si la acción no es pública y `body.token` no coincide con el hash guardado en Script Properties (`ADMIN_HASH`), responde `{ ok:false, error:"No autorizado" }`. Si `ADMIN_HASH` no está configurado, las acciones admin se bloquean (seguro por defecto).
- **`validarPin`** (pública): recibe `token` (SHA-256 hex del PIN) y devuelve `{ ok:true, valido:true|false }`. La usa la pantalla de PIN.
- **`configurarPin()`**: función de setup que se ejecuta una sola vez desde el editor de Apps Script. Mateo escribe el PIN nuevo en una constante, la ejecuta (guarda el hash SHA-256 hex en Script Properties) y vuelve a dejar el placeholder. Incluye helper `_sha256hex()` con `Utilities.computeDigest` (hex minúsculas, UTF-8 — debe coincidir con `crypto.subtle` del frontend).

### Frontend (`index.html`)

- Se elimina el hash hardcodeado `_PH`.
- `checkPin()`: hashea el input y llama a `api({action:"validarPin", token:hash})`. Si `valido`, guarda `ma_auth=1` y `ma_tok=hash` en sessionStorage y arranca la app. Distingue "PIN incorrecto" de error de red. Muestra "Verificando…" mientras espera.
- `api()`: si existe `sessionStorage.ma_tok`, lo agrega como `token` al body de todos los requests. Como todas las llamadas admin pasan por `api()`, no hay más cambios.
- Gate de inicio (línea ~4142): saltea la pantalla de PIN solo si existen `ma_auth` **y** `ma_tok` (una sesión vieja sin token debe volver a pedir PIN).
- `formulario_inscripcion.html` y `s/index.html`: sin cambios (usan solo acciones públicas).

### Seguridad

El hash ya no aparece en ningún código fuente servido; solo vive en Script Properties. El único ataque posible es fuerza bruta online contra `validarPin` (una llamada HTTP por intento) — inviable con 8+ dígitos y las cuotas de Apps Script.

## Mejora 3 — `enviarMails` robusto

Todo en `sorteo_script.gs` (función `enviarMails`, líneas ~270-356):

1. **Lock de concurrencia**: la función se envuelve en `LockService.getScriptLock()` (mismo patrón que `upsertShow`). Si no obtiene el lock en 30s devuelve error "Otro envío está en curso". Evita que 2 usuarias enviando a la vez dupliquen mails.
2. **Pre-check de PDFs en el servidor**: después de juntar los ganadores pendientes, si `pdfs.length < ganadores.length × entradasXGan`, devuelve `{ ok:false, error:"Hay X PDF(s) y se necesitan Y — no se envió ningún mail." }` sin enviar nada.
3. **Marcado Enviando → Enviado**: por cada ganador, la fila pasa a "Enviando" *antes* de `GmailApp.sendEmail`, y a "Enviado" después. Si el envío falla, vuelve a "Pendiente" (solo si el mail no salió — flag `sent`). Un reintento solo toma filas "Pendiente", así que un corte a mitad de ejecución nunca duplica mails (a lo sumo deja filas en "Enviando" para revisar a mano).
4. **Respuesta más precisa**: devuelve `enviadosMails` (lista de mails efectivamente enviados). Comparaciones de estado con `String(...).trim()`.

### Frontend

- `_enviarMailsCore`: marca como "enviado" localmente solo los ganadores cuyo mail está en `res.enviadosMails` (hoy marca todos aunque algunos hayan fallado, lo que rompe el reintento desde la UI).
- `mostrarPreflightPDFs`: los estados `insufficient` y `no-show-folder` pierden el botón "Enviar igual" (el servidor lo rechazaría); queda solo "Cancelar" con la explicación.

## Deploy (orden importa)

1. Pegar el nuevo `sorteo_script.gs` en el editor de Apps Script.
2. Ejecutar `configurarPin()` con el PIN nuevo (y restaurar el placeholder).
3. **Actualizar la implementación existente** a una versión nueva (Implementar → Administrar implementaciones → ✏️ → Versión nueva). NO crear implementación nueva (cambiaría la URL). El modo queda "Yo + Cualquier persona".
4. Publicar el nuevo `index.html` (GitHub Pages + Netlify).
5. Avisar el PIN nuevo a la otra usuaria.

## Plan de prueba manual (~5 min)

1. PIN incorrecto → "PIN incorrecto". PIN correcto → entra y el dashboard sincroniza (token funciona).
2. Desde una pestaña privada, POST directo a la URL del script con `{action:"getInscriptos", showId:"X"}` sin token → debe responder "No autorizado". Con `{action:"validarMail", mail:"..."}` → debe seguir funcionando (formulario no se rompe).
3. Show de prueba con 2 ganadores × 1 entrada y carpeta de Drive con 1 solo PDF → "Enviar via Gmail" debe fallar con el mensaje de PDFs insuficientes y ningún mail enviado.
4. Completar los PDFs → enviar → llegan los mails, filas en "Enviado". Tocar "Enviar via Gmail" de nuevo → "No hay ganadores pendientes" (sin duplicados).

## Fuera de alcance

- Lock en `trackingGanadores` (hallazgo #2 del análisis — pendiente para otra iteración).
- Privacidad de `leerTracking`/`getInscriptos` más allá del token (con token ya quedan protegidas).
- Limpieza de archivos viejos (`files/`, PNGs, funciones de debug).
