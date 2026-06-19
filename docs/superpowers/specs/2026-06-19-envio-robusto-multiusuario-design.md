# Envío robusto multiusuario: cero-pendientes claro + clave de envío consistente — Design

**Fecha:** 2026-06-19
**Defectos:** #3 y #2 del diagnóstico del bug de envío (memoria `project-bug-envio-angela-leiva`).
**Relación:** continúa el bloque 1 (confirmación durable, PR #1). Es independiente: se puede mergear en cualquier orden respecto del bloque 1.
**Orden de implementación:** primero defecto 3 (seguro, hace visible el problema), después defecto 2 (lo arregla). El defecto 3 queda como red de seguridad permanente.

---

## Defecto 3 — distinguir "nada que enviar" de "desincronización"

### Problema

Al apretar "Enviar via Gmail" sin nada para enviar, el sistema muestra **siempre** el mismo mensaje gris ambiguo ("No hay pendientes"), sin distinguir tres situaciones:

- **(a)** Ya se enviaron las entradas a todos los ganadores del show.
- **(b)** El show todavía no tiene ganadores sorteados.
- **(c)** Hay ganadores que deberían poder enviarse, pero la planilla devuelve cero pendientes (desincronización — el escenario del bug de Angela Leiva).

El caso (c) es el peligroso: hoy se ve igual que (a).

### Comportamiento por caso (todo frontend)

El frontend conoce los conteos locales del evento activo (`S.ganadores` filtrado por `evId`).

**(a)/(b) — en `enviarGmail` ([index.html:3774-3778](index.html))**. El branch actual:

```javascript
  const pend=S.ganadores.filter(g=>g.evId===evId&&g.estado==="pendiente");
  if(!pend.length){toast("No hay pendientes","warn");return;}
```

se reemplaza por una distinción según haya o no ganadores locales del evento:

- 0 ganadores locales → **(b)** neutro: *"Este show todavía no tiene ganadores sorteados."*
- Hay ganadores pero ninguno pendiente (todos `enviado`) → **(a)** éxito: *"Ya enviaste las entradas a los N ganadores de este show."* (N = total de ganadores del evento)

**(c) — en `_enviarMailsCore` ([index.html:3798-3808](index.html))**. Solo se llama con `pend.length>0`. Tras la respuesta del backend, además del guard `if(!res.ok)` existente, se agrega antes del toast de éxito:

- Si `res.enviados===0` **y** no hay `res.errores` **y** `pend.length>0` → desincronización. Mostrar **alerta** y **no** marcar nada como enviado, luego `return`:
  *"Esperaba N ganador(es) pendiente(s) para este show, pero la planilla no devolvió ninguno. Posible desincronización — revisá/reintentá la sincronización antes de enviar."* (N = `pend.length`)

### Tabla de decisión (frontend)

| Situación | `pend` local | total ganadores evento | `res.enviados` | `res.errores` | Resultado |
|---|---|---|---|---|---|
| (b) sin ganadores | 0 | 0 | — (no llama backend) | — | Neutro: "no tiene ganadores sorteados" |
| (a) todos enviados | 0 | >0 | — (no llama backend) | — | Éxito: "ya enviaste a los N" |
| envío normal | >0 | >0 | >0 | — | Flujo actual sin cambios |
| errores de envío | >0 | >0 | 0 o parcial | >0 | `alert` de errores (existente) |
| (c) desincronización | >0 | >0 | 0 | vacío | **Alerta**: "esperaba N, planilla devolvió 0" |

### Backend (opcional, mínimo)

El `mensaje` del caso "0 pendientes" ([sorteo_script.gs:340](sorteo_script.gs)) puede hacerse más explícito, **sin** cambiar `ok` ni la estructura. Opcional: el frontend ya distingue los tres casos solo.

---

## Defecto 2 — clave de envío consistente (envío multiusuario)

### Problema

"Desde otra computadora no se envían los mails" (y, en general, después de recargar un show desde la planilla). Causa raíz, confirmada con auditoría de las planillas:

- La **escritura** de ganadores usa `ev.id` como `showId` ([index.html:3582](index.html)).
- El **envío** usa `ev.nombre` (`nombreCarpeta`) como `showId` ([index.html:3779](index.html),[:3801](index.html)).
- El backend matchea filas de Ganadores por `showId` **exacto** ([sorteo_script.gs:328](sorteo_script.gs)).

Cuando `ev.id ≠ ev.nombre`, el envío no encuentra las filas → 0 enviados, sin error. La auditoría halló **46 de 71 shows con ID ≠ Nombre** (patrón `ARTISTA_DD-MM-YYYY` como ID vs `ARTISTA - DD/MM/YYYY` como Nombre, más nombres pelados con prefijo de fecha en el Nombre). Como `ev.id` es el ID canónico de la hoja Shows —igual en todas las compus vía `getShowsCloud` ([index.html:3911](index.html))— alinear el envío a `ev.id` resuelve el problema sin tocar datos.

### Enfoque elegido: A + limpieza puntual

Usar `ev.id` (clave canónica de Shows) en envío y búsqueda de PDFs, en vez de `ev.nombre`. **Sin migración de datos.** Más una limpieza manual de la entrada duplicada de Angela Leiva.

### Cambios (todo frontend)

1. **Envío** ([index.html:3801](index.html), dentro de `_enviarMailsCore`): llamar `enviarMails` con `showId: ev.id`. Así matchea las filas de Ganadores (que ya se guardaron con `ev.id`).
2. **Búsqueda de PDFs** ([index.html:3782-3788](index.html), `enviarGmail` → `checkPDFs`): pasar `showId: ev.id` **y** `showNombre: ev.nombre`. El backend matchea la carpeta de Drive de forma flexible (por nombre, por id, o substring case-insensitive — [sorteo_script.gs:566-588](sorteo_script.gs)), así que las carpetas existentes se siguen encontrando.
3. **Threading del identificador**: `enviarGmail` hoy deriva `nombreCarpeta=ev?.nombre||evId` y lo pasa a `mostrarPreflightPDFs(check,pend,ev,nombreCarpeta)` y `_enviarMailsCore(nombreCarpeta,ev,pend)`. Se refactoriza para que la **clave de envío** sea `ev.id` y el **nombre para Drive/UI** sea `ev.nombre`, sin perder el fallback a `evId` cuando falten.
4. **`ev.id` canónico y único**: verificar que el evento usado para sortear/enviar provenga de la hoja Shows (cloud), no de un evento "fantasma" reconstruido por `importarTracking` con id en formato `SHOW_fecha` ([index.html:2489-2502](index.html)). Si `importarTracking` puede crear un evento que compita con el del cloud para un show activo, ajustarlo para que no pise el `ev.id` canónico (p. ej. preferir match contra eventos ya existentes del cloud antes de crear `_fromTracking`).

### Limpieza puntual (operación manual, no parte del código)

En la hoja **Shows**: unificar/eliminar la entrada duplicada de Angela Leiva (`ANGELA LEIVA` activa, cupo 15, vs `2026-06-18 ANGELA LEIVA` ya marcada Eliminado). Confirmar con el usuario cuál conservar. No es una migración masiva.

### Por qué funciona multiusuario

A y B cargan el mismo `ev.id` desde la hoja Shows; escriben y envían con ese mismo id → siempre coinciden. No depende de `ev.nombre`, que era el campo que divergía.

### Compatibilidad con históricos

Los shows ya enviados con `showId` en otro formato (JONAS BROTHERS, QLOKURA, etc.) quedan como están — no se re-envían. Los pendientes actuales (ARJONA) matchean porque su `showId` de Ganadores ya coincide con un ID de Shows.

---

## Testing

Sin framework de tests (HTML + Apps Script). Verificación manual en navegador + verificación estática (`node --check` del bloque `<script>`).

**Defecto 3** — reproducir los tres casos:
- (a) evento con todos los ganadores en `estado:"enviado"`.
- (b) evento recién creado, sin sortear.
- (c) stub de `window.api` que hace que `enviarMails` devuelva `{ok:true, enviados:0, enviadosMails:[], errores:[], mensaje:"No hay ganadores pendientes"}` con `pend` local > 0.

**Defecto 2** — verificar que tras la corrección, el `showId` enviado a `enviarMails`/`checkPDFs` es `ev.id`; simular el escenario multi-compu: cargar un show desde la nube cuyo `ID ≠ Nombre` (p. ej. uno con formato `ARTISTA_DD-MM-YYYY`) y confirmar que el envío matchea las filas de Ganadores guardadas con ese `ev.id`.

## Error handling

- Caso (c) no marca ganadores como enviados ni modifica `S.ganadores`/`ticketsBase`.
- Guards existentes (`if(!res.ok)`, `alert` de `res.errores`) se conservan.
- El cambio de clave de envío no altera el flujo cuando sí hay pendientes y los ids coinciden.

## Fuera de alcance

- Migración masiva de IDs históricos (enfoque B descartado).
- Slug opaco nuevo.
- Cambios al flujo de envío cuando sí hay pendientes y todo coincide.
