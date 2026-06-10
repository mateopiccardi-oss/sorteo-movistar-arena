# Concurrencia entre usuarios al editar eventos

**Fecha:** 2026-06-10
**Autor:** Mateo Piccardi
**Contexto:** Sistema de Sorteos — Movistar Arena RRHH

## Problema

El sistema lo operan 2 personas de RRHH en paralelo, cada una desde su navegador. Hoy:

1. La pestaña `Shows` del Sheet ya guarda la metadata de los eventos (`upsertShow`/`getShowsCloud`), pero el sync solo ocurre al abrir la app (`syncShowsFromCloud` a los ~800 ms del arranque).
2. `upsertShow` se llama fire-and-forget desde el frontend, sin esperar respuesta.
3. No hay detección de conflicto: si dos usuarios editan el mismo evento sin recargar, el último pisa al primero en silencio.

El choque ocurre poco, pero cuando ocurre se pierden cambios sin aviso. La solución tiene que ser sutil — sin lock duro, sin refresh automático intrusivo, sin auditoría visible.

## Objetivos

- Que ningún guardado pise cambios ajenos sin avisar.
- Que el operador se entere de cambios remotos en eventos relevantes sin que la app lo interrumpa.
- Mantener cero build step, cero dependencias nuevas.
- Cero costo perceptible en cuota de Apps Script (la app vive del límite gratuito).

## No-objetivos

- Lock duro al abrir el editor.
- Refresh en vivo del editor mientras el usuario tipea.
- Mostrar quién editó qué (auditoría) — descartado por YAGNI.
- Resolver concurrencia en el flujo de sorteo/ganadores/mails (otro problema, otro diseño).
- Resolver creación simultánea del mismo evento con `id` idéntico (caso prácticamente nulo porque `id` se deriva del nombre del show).

## Diseño

### 1. Versionado optimista por evento (backend)

La columna `ActualizadoEn` de la hoja `Shows` (col 12, índice 11) actúa como versión del evento. Es escrita por el servidor (Apps Script), no por el cliente, así que no depende del reloj local.

`upsertShow(show, expectedVersion?)` acepta un parámetro opcional nuevo:

- **Sin `expectedVersion`** → comportamiento actual: sobrescribe. Se usa al crear un evento, al "pisar igual" tras un conflicto consciente, y al restaurar un evento eliminado.
- **Con `expectedVersion`** → antes de escribir, lee el `ActualizadoEn` actual de la fila:
  - Coincide → guarda, devuelve `{ ok: true, newVersion: "<dd/MM/yyyy HH:mm:ss>" }`.
  - No coincide → no toca nada, devuelve `{ ok: false, error: "conflict", currentVersion: "<dd/MM/yyyy HH:mm:ss>", currentShow: <objeto del row> }`.

`getShowsCloud` agrega `actualizadoEn` al objeto que devuelve (1 línea — col 12, índice 11). `creadoEn` ya se devuelve.

`deleteShow` se mantiene como está (sin guarda de versión). El poller (sección 3) detecta el delete y lo refleja en el frontend.

### 2. Versión en el estado del frontend (S.eventos)

Cada evento en `S.eventos` gana dos campos relacionados pero distintos:

- `_version` — string, la última versión cloud que el cliente conoce de ese evento. Se actualiza cada vez que llega un `actualizadoEn` nuevo del cloud (incluido cuando se muestra el banner). Su rol: evitar que el poller dispare el banner una y otra vez para el mismo cambio.
- `_versionAtOpen` — string, snapshot del `_version` en el momento exacto en que se abre el editor. **No** se actualiza cuando llega un cambio del cloud. Su rol: ser el `expectedVersion` que viaja al backend, representando "esta es la versión sobre la que yo empecé a editar".

Distinción clave: `_version` rastrea lo que está en la nube; `_versionAtOpen` rastrea desde dónde empecé a editar. Pueden divergir mientras el editor está abierto.

Al guardar:

- Si el evento ya existía: `upsertShow({ ...ev, expectedVersion: _versionAtOpen })`.
- Si es un evento recién creado: `upsertShow(ev)` sin `expectedVersion`.

Manejo del `{ error: "conflict" }`:

- Modal: *"Este evento cambió en la nube mientras lo editabas (última actualización: hace X). ¿Qué querés hacer?"*
- Botones: **[Ver cambios primero]** (recarga el editor con `currentShow`, descarta los cambios locales con confirmación) — **[Pisar igual]** (re-llama `upsertShow` sin `expectedVersion`).

### 3. Poller silencioso (pollShowsLight)

Función nueva `pollShowsLight()` que se ejecuta cada 60 s:

- **Gate de visibilidad:** corre solo si `document.visibilityState === "visible"`. En tabs en segundo plano no consume cuota.
- **Anti-overlap:** flag `_pollInFlight`; si ya hay una request en vuelo, se descarta el tick.
- **Errores silenciosos:** un fetch que falla se loguea por consola y reintenta en el próximo tick. Nunca aparece toast de error.

Algoritmo:

1. Llama `getShowsCloud()`.
2. Por cada show del cloud, busca el evento local por `id`.
3. **Evento nuevo en cloud, no existe local** → agregar a `S.eventos` con su `_version`.
4. **Existe local y `actualizadoEn` del cloud > `_version` local:**
   - Si **no** es el evento abierto en el editor **y no** es `S.sorteo.evId` → reemplaza los campos del evento en `S.eventos` con los del cloud, actualiza `_version`, re-renderiza el dashboard.
   - Si **es** el evento abierto en el editor → no se tocan los campos del evento en `S.eventos` (para no pisar nada que el usuario tenga en pantalla), pero **sí** se actualiza `_version` a la versión nueva del cloud y se dispara el banner (sección 4). `_versionAtOpen` queda intacto.
   - Si **es** `S.sorteo.evId` (sorteo en curso) → no se toca. El sync queda diferido hasta que termine el sorteo.
5. **Evento existe local pero no está en el cloud** (soft-delete) → quitar de `S.eventos`, salvo que esté abierto en el editor o sea `S.sorteo.evId`. Si está abierto, disparar el banner de delete (sección 4).
6. Persistir con `save()` y re-renderizar dashboard si hubo cambios.

### 4. Banner contextual en el editor

Elemento nuevo arriba del formulario de edición de evento. Estilos en línea con la paleta amber existente (`var(--amber)`/`var(--amber-bg)`). Tres estados posibles:

**Estado A — Evento actualizado por otra persona:**

> *"Otra persona actualizó este evento hace 30 s. **[Ver cambios]** **[Seguir editando]**"*

- **Ver cambios** → si hay campos del form modificados localmente, pedir confirmación; recargar los inputs del editor con los valores del cloud y reasignar `_versionAtOpen = _version` (ahora estoy editando sobre la versión nueva).
- **Seguir editando** → ocultar el banner. `_versionAtOpen` no se toca. Si después el usuario aprieta Guardar, salta el modal de conflicto (sección 2) porque `_versionAtOpen` sigue siendo viejo respecto al cloud.

**Estado B — Evento eliminado por otra persona:**

> *"Este evento fue eliminado por otra persona. **[Restaurar y guardar mis cambios]** **[Descartar]**"*

- **Restaurar** → `upsertShow(ev)` sin `expectedVersion`. El evento vuelve a aparecer.
- **Descartar** → cerrar el editor, quitar el evento de `S.eventos`.

Estado A y B son mutuamente excluyentes — si el evento se borró, eso pisa el banner de "actualizado".

### 5. Inicialización

En `startApp()`, después del `setTimeout(syncShowsFromCloud, 800)` actual, se arma el poller con `setInterval(pollShowsLight, 60000)`. La primera ejecución del poller queda 60 s después del arranque (el primer sync ya lo hace `syncShowsFromCloud`).

No hace falta limpiar el interval — la pestaña vive lo que vive el operador.

## Data flow

```
Usuario A abre editor del evento X
  ev._versionAtOpen = "10/06/2026 14:30:00"   ← snapshot

Usuario B (en otro browser) edita X y guarda
  upsertShow(X, expectedVersion="10/06/2026 14:30:00")
  → ok: true, newVersion="10/06/2026 14:32:15"

Poller de A (60s después) detecta cambio
  cloudShow.actualizadoEn = "10/06/2026 14:32:15"
  local.ev._version = "10/06/2026 14:30:00"
  cloud > local → evento abierto → banner Estado A

Usuario A elige "Seguir editando" y aprieta Guardar
  upsertShow(X, expectedVersion="10/06/2026 14:30:00")
  → ok: false, error: "conflict", currentVersion="10/06/2026 14:32:15"
  Modal de conflicto

Usuario A elige "Pisar igual"
  upsertShow(X)   ← sin expectedVersion
  → ok: true, newVersion="10/06/2026 14:34:50"
```

## Edge cases

| Caso | Comportamiento |
|---|---|
| Poller falla por red | Silencio. Retry en 60 s. |
| Apps Script tarda > 60 s en responder | Flag `_pollInFlight` descarta el tick siguiente hasta que vuelva. |
| Reloj del cliente desincronizado | Irrelevante — `actualizadoEn` lo escribe el servidor. |
| Dos creaciones simultáneas del mismo `id` | El segundo pisa al primero (sin guarda). Caso prácticamente nulo: `id` se deriva del nombre del show. |
| Evento abierto en editor + cloud delete | Banner Estado B. |
| Sorteo en curso (`S.sorteo.evId`) | Ese evento no se sincroniza visualmente hasta que termine el sorteo. |
| Usuario A "Pisa igual" después de conflicto | Funciona — el `upsertShow` sin `expectedVersion` es el escape hatch consciente. |
| Pestaña en background varias horas | El poller no corre (visibilityState). Al volver al foreground, el próximo tick (≤60 s después) sincroniza todo. |
| `getShowsCloud` devuelve `ok:false` | El poller no hace nada ese tick. Como `syncShowsFromCloud` ya cubre el arranque, no se pierde nada. |

## Cuota Apps Script

60 s × 2 operadores × 8 h/día = ~960 calls/día por `getShowsCloud`. Apps Script da 20.000 calls/día gratis. Sobra ampliamente. El `upsertShow` con verificación de versión cuesta una lectura extra de la hoja, despreciable.

## Cambios concretos por archivo

**`sorteo_script.gs`:**
- `upsertShow`: aceptar `show.expectedVersion` opcional, leer `ActualizadoEn` de la fila si hay match de `id`, devolver `conflict` si difiere. Devolver `newVersion` en éxito.
- `getShowsCloud`: agregar `actualizadoEn: String(datos[i][11] || "").trim()` al objeto del show.

**`index.html`:**
- Al recibir shows del cloud (`syncShowsFromCloud` y poller): copiar `s.actualizadoEn` a `ev._version`.
- Al abrir editor: setear `_versionAtOpen` desde el `_version` del evento.
- En el handler de guardar (líneas ~1402, ~1448 según grep): pasar `expectedVersion: ev._versionAtOpen` en la llamada `upsertShow`, y dejar de ser fire-and-forget — `await` el resultado.
- Manejo de `error: "conflict"`: modal con dos botones.
- Función nueva `pollShowsLight()` + `setInterval` en `startApp()`.
- Elemento HTML nuevo del banner dentro del editor de eventos + 2 funciones de show/hide y handlers.

**Sin tocar:**
- `formulario_inscripcion.html` (no edita eventos).
- Lógica de inscripciones, sorteo, ganadores, mails, Tracking.
- `deleteShow`.

## Testing manual (no hay tests automáticos)

Lista de comprobaciones a correr antes de mergear:

1. Crear evento nuevo desde browser A → aparece en B en ≤60 s sin recargar.
2. Editar `cantidad` desde A, guardar → aparece en B en ≤60 s.
3. Eliminar desde A → desaparece en B en ≤60 s.
4. A y B abren el editor del mismo evento. A guarda. B guarda → B ve modal de conflicto.
5. Mismo caso anterior, B elige "Ver cambios" → form se recarga con los datos de A, sin error.
6. Mismo caso, B elige "Pisar igual" → guarda OK, A ve banner Estado A en ≤60 s.
7. A elimina el evento. B lo tiene abierto en el editor → B ve banner Estado B.
8. B elige "Restaurar y guardar mis cambios" → evento vuelve a aparecer en A en ≤60 s.
9. Pestaña en background 5 min mientras A hace cambios → al volver al foreground, sincroniza en ≤60 s.
10. Apagar wifi 2 min, hacer cambios locales, reconectar → poller no rompe nada, próximo guardado funciona o tira conflict según corresponda.
11. Iniciar sorteo en A. B modifica el evento en sorteo → A no ve el cambio hasta terminar el sorteo (sin banner).
