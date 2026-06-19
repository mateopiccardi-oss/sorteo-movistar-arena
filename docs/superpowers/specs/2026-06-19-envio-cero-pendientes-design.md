# Envío con cero pendientes: distinguir "nada que enviar" de "desincronización" — Design

**Fecha:** 2026-06-19
**Defecto:** #3 del diagnóstico del bug de envío (memoria `project-bug-envio-angela-leiva`).
**Alcance:** solo este defecto. El defecto #2 (ID = nombre inestable / slug estable + migración) queda para un plan aparte.

## Problema

Cuando se aprieta "Enviar via Gmail" y no hay nada para enviar, el sistema muestra **siempre el mismo** mensaje gris ambiguo ("No hay pendientes"), sin distinguir entre tres situaciones muy distintas:

- **(a)** Ya se enviaron las entradas a todos los ganadores del show.
- **(b)** El show todavía no tiene ganadores sorteados.
- **(c)** Hay ganadores que deberían poder enviarse, pero la planilla devuelve cero pendientes (desincronización — el escenario del bug de Angela Leiva).

El caso (c) es el peligroso: hoy se ve igual que "ya está todo enviado", así que una falla de guardado pasa desapercibida.

## Objetivo

Que el botón "Enviar" comunique **cuál de los tres casos** está ocurriendo, para que (c) salte a la vista. No cambia nada del flujo cuando sí hay pendientes para enviar.

## Arquitectura

Todo el cambio de decisión vive en el **frontend** (`index.html`), que ya conoce los conteos locales del evento activo (`S.ganadores` filtrado por `evId`). No hay migración de datos. El backend queda sin cambios funcionales (solo, opcionalmente, un mensaje más claro que NO altera `ok`/estructura).

Es **independiente del bloque 1** (PR de confirmación durable): se puede mergear en cualquier orden. Si el bloque 1 ya está, el banner de "sin sincronizar" complementa naturalmente el caso (c); si no, el caso (c) igual funciona con su propio mensaje. El caso (c) NO depende del flag `_synced`.

## Comportamiento por caso

### (a) Todos ya enviados / (b) sin ganadores — en `enviarGmail` (`index.html:3774-3778`)

El branch actual:

```javascript
  const pend=S.ganadores.filter(g=>g.evId===evId&&g.estado==="pendiente");
  if(!pend.length){toast("No hay pendientes","warn");return;}
```

se reemplaza por una distinción según haya o no ganadores locales del evento:

- Si el evento tiene **0 ganadores** locales → caso (b), mensaje neutro:
  *"Este show todavía no tiene ganadores sorteados."*
- Si tiene ganadores pero **ninguno pendiente** (todos `enviado`) → caso (a), mensaje de éxito/neutro:
  *"Ya enviaste las entradas a los N ganadores de este show."* (N = total de ganadores del evento)

El conteo se obtiene con `S.ganadores.filter(g=>g.evId===evId)`.

### (c) Desincronización — en `_enviarMailsCore` (`index.html:3798-3808`)

`_enviarMailsCore` solo se llama cuando `pend.length>0` (ya pasó el guard de `enviarGmail`). Tras la respuesta del backend, además del guard `if(!res.ok)` existente, se agrega la detección del caso (c) **antes** del toast de éxito:

- Si `res.enviados===0` **y** no hay `res.errores` (no fueron errores de envío) **y** `pend.length>0` (esperábamos mandar algo) → es desincronización. Mostrar **alerta** y **no** marcar nada como enviado:
  *"Esperaba N ganador(es) pendiente(s) para este show, pero la planilla no devolvió ninguno. Posible desincronización — revisá/reintentá la sincronización antes de enviar."* (N = `pend.length`)
  Luego `return` (no cae en el toast de éxito ni en el marcado local).

Esto distingue (c) de:
- envío exitoso (`res.enviados>0`) → flujo normal sin cambios;
- errores de envío (`res.errores.length>0`) → ya manejado por el `alert` de errores existente (`index.html:3808`).

### Backend (opcional, mínimo) — `sorteo_script.gs:340`

El mensaje del caso "0 pendientes" puede hacerse más explícito, pero **sin** cambiar `ok` ni la estructura del objeto (otros llamadores dependen de eso). Es opcional porque el frontend ya distingue los tres casos por su cuenta. Si se hace, es solo texto del campo `mensaje`.

## Estados / tabla de decisión (frontend)

| Situación | `pend` local | total ganadores del evento | `res.enviados` | `res.errores` | Resultado |
|---|---|---|---|---|---|
| (b) sin ganadores | 0 | 0 | — (no llama backend) | — | Neutro: "no tiene ganadores sorteados" |
| (a) todos enviados | 0 | >0 | — (no llama backend) | — | Éxito: "ya enviaste a los N" |
| envío normal | >0 | >0 | >0 | — | Flujo actual sin cambios |
| errores de envío | >0 | >0 | 0 o parcial | >0 | `alert` de errores (existente) |
| (c) desincronización | >0 | >0 | 0 | vacío | **Alerta roja**: "esperaba N, la planilla devolvió 0 — posible desincronización" |

## Error handling

- El caso (c) NO marca ganadores como enviados y NO modifica `S.ganadores` ni `ticketsBase`.
- Los guards existentes (`if(!res.ok)`, `alert` de `res.errores`) se conservan intactos.

## Testing

Sin framework de tests (HTML + Apps Script). Verificación manual en navegador, reproduciendo los tres casos:
- (a) un evento con todos los ganadores en `estado:"enviado"`.
- (b) un evento recién creado, sin sortear.
- (c) stub que hace que el backend devuelva `{ok:true, enviados:0, enviadosMails:[], errores:[], mensaje:"No hay ganadores pendientes"}` mientras hay `pend` local > 0 (mismo patrón de stub de `window.api` usado en el bloque 1).
Más verificación estática (`node --check` del bloque `<script>`).

## Fuera de alcance

- Defecto #2 (ID = nombre libre → slug estable único + migración de IDs en Shows/Ganadores).
- Cualquier cambio al flujo de envío cuando sí hay pendientes.
