# Etapa 1 — Base de almacenamiento confiable

**Fecha:** 2026-06-29
**Estado:** Diseño aprobado (pendiente revisión del spec por el usuario)
**Alcance:** Solo frontend (`index.html`). NO toca el Apps Script → sin riesgo de CORS/deploy.

## Contexto

El proyecto de "envío confiable" se ataca en 3 etapas, en orden, cada una probada antes de seguir:

1. **Etapa 1 (este spec):** que el navegador no se llene y `save()` nunca falle en silencio.
2. **Etapa 2 (futuro):** confirmación durable de ganadores (que siempre queden en la planilla o se avise fuerte).
3. **Etapa 3 (futuro):** envío en tandas sin cortarse + leer el estado real desde la planilla.

Este documento cubre **solo la Etapa 1**.

## Problema

La función `save()` ([index.html:1265](../../../index.html)) persiste en `localStorage` (clave `ma_v6`) **todos** los ganadores, incluidos los históricos reconstruidos desde la planilla. Hoy hay ~959 y suben con cada sorteo. `localStorage` tiene un límite (~5 MB); al superarlo, `setItem` lanza `QuotaExceededError`, que el `catch` actual solo registra con `console.warn` ([index.html:1281](../../../index.html)) — **falla en silencio**.

Consecuencia: cuando se llena, deja de guardarse **todo** (pendientes, estados de envío, confirmaciones), no solo los históricos. Esto explica el bug "el banner no sobrevive a la recarga" del PR #1 y contribuye a las desincronizaciones de envío (caso Fito Páez, 2026-06-29).

Ver memoria del proyecto: bug de envío Angela Leiva / recurrencia Fito Páez.

## Hecho clave que habilita la solución

Los ganadores históricos **ya se cargan desde la planilla automáticamente al abrir la app**: `setTimeout(() => importarDesdeTracking(true), 1500)` ([index.html:4136](../../../index.html)). `importarDesdeTracking` ([index.html:2454](../../../index.html)) reconstruye `S.ganadores` desde la hoja "Tracking Ganadores", marcando cada uno con `_historico: true`. El anti-repetidos (`ganoPrevio` / `ganoPrevioInfo`, [index.html:2952](../../../index.html)) los consume de `S.ganadores`.

Por lo tanto, **la planilla ya es la fuente de verdad de los históricos**; guardarlos también en `localStorage` es duplicación innecesaria.

## Decisión de diseño

Enfoque elegido: **A — no persistir históricos en el navegador + detectar fallo de guardado** (descartados: B comprimir+purgar, C migrar a IndexedDB, por complejidad y por no atacar la raíz).

Para vos en pantalla **no cambia nada**: mismo historial, mismo conteo, mismas exclusiones "Ya ganó 🏆". Solo dejan de duplicarse en el navegador.

## Cambios

Tres cambios, todos en `index.html`.

### Cambio 1 — `save()` no guarda históricos

En `save()` ([index.html:1265](../../../index.html)), al armar `toSave`, filtrar:

```js
ganadores: S.ganadores.filter(g => !g._historico),
```

- Solo se persisten los ganadores **no históricos** (del show activo: pendientes / recién confirmados / enviados de la sesión).
- **Migración automática:** el primer `save()` con el filtro reescribe `ma_v6` sin los ~959 históricos. No requiere acción manual ni script de migración.
- `load()` ([index.html:1284](../../../index.html)) no necesita cambios: al abrir habrá pocos (o ningún) ganador en `ma_v6`, e `importarDesdeTracking(true)` repuebla los históricos a los 1,5 s, como ya hace hoy.

### Cambio 2 — Aviso visible si `save()` falla

Hoy el `catch` de `save()` solo hace `console.warn`. Cambiar a:

- Marcar un estado de "guardado fallido" (variable a nivel de módulo, p. ej. `_saveError`).
- Mostrar un **cartel rojo persistente** en la parte superior: *"⚠ No se pudo guardar localmente. No cierres la pestaña."* Distinguir `QuotaExceededError` de otros errores solo a nivel de log (el mensaje al usuario es el mismo).
- Cuando un `save()` posterior tenga éxito, limpiar el estado y ocultar el cartel.

Con la Etapa 1 esto casi no debería dispararse (el navegador queda liviano), pero queda como red de seguridad permanente.

### Cambio 3 — Bloquear el sorteo si los históricos no cargaron

Como los históricos ahora dependen de la carga desde la planilla, si esa carga falla (p. ej. sin internet) el anti-repetidos quedaría ciego. Decisión del usuario: **bloquear el sorteo y avisar** (no dejar sortear a ciegas).

- Agregar una marca de estado, p. ej. `_historicosCargados` (boolean), inicialmente `false`.
- `importarDesdeTracking` la pone en `true` al completar OK; en `false` (o la deja en `false`) si falla. Hoy corre con `silente=true` y se traga los errores — hay que capturar el resultado para setear la marca, sin romper el comportamiento silencioso para los toasts.
- En `ejecutarSorteo()` ([index.html:3046](../../../index.html)): si `_historicosCargados !== true`, **no iniciar el sorteo**; mostrar un cartel/aviso: *"No se pudieron cargar los ganadores históricos — el anti-repetidos no funciona. Reintentá."* con un botón **Reintentar** que vuelve a llamar a `importarDesdeTracking(false)`.
- El botón `#btn-sort` ([index.html:827](../../../index.html)) se muestra deshabilitado mientras `_historicosCargados !== true`. Apenas la carga termina OK, se habilita.
- Caso `antiRep === false` (evento sin anti-repetición): el bloqueo **no aplica** — si el evento no usa anti-repetidos, no importa que los históricos no estén; dejar sortear.

## Flujo de datos resultante

1. **Abrir app:** `load()` lee de `ma_v6` solo no-históricos (pocos). `#btn-sort` deshabilitado.
2. **~1,5 s después:** `importarDesdeTracking(true)` trae históricos de la planilla → `_historicosCargados = true` → `#btn-sort` habilitado. `S.ganadores` = no-históricos (local) + históricos (planilla, en memoria).
3. **Cualquier cambio:** `save()` persiste solo no-históricos. Históricos nunca tocan el navegador.
4. **Si `importarDesdeTracking` falla:** `_historicosCargados = false` → sorteo bloqueado + cartel con Reintentar.
5. **Si `save()` falla:** cartel rojo persistente hasta el próximo guardado exitoso.

## Casos borde

- **Evento sin anti-repetición (`antiRep === false`):** no se bloquea el sorteo aunque falten históricos.
- **Históricos que aún no se sincronizaron a la planilla:** fuera de alcance de la Etapa 1 (lo cubre la Etapa 2, confirmación durable). El comentario en `importarDesdeTracking` ([index.html:2469-2471](../../../index.html)) asume que los no-sincronizados ya están marcados `_historico` tras un sync OK; esa fragilidad se aborda en la Etapa 2.
- **Multi-usuario ([[project_dos_usuarios]]):** sin cambios de comportamiento respecto a hoy; cada navegador carga históricos de la misma planilla compartida.

## Criterios de éxito

1. Tras un sorteo y varios envíos, `localStorage["ma_v6"]` **no** contiene ganadores con `_historico: true`.
2. El tamaño de `ma_v6` se mantiene chico (< ~100 KB en uso normal), independientemente de cuántos shows históricos haya.
3. El historial, el conteo y las exclusiones "Ya ganó 🏆" se ven y funcionan igual que antes (datos desde la planilla).
4. Si se fuerza un fallo de `save()`, aparece el cartel rojo y desaparece tras un guardado exitoso.
5. Si los históricos no cargan, el botón de sortear queda deshabilitado con el cartel de reintento (salvo eventos con `antiRep === false`).

## Fuera de alcance (etapas siguientes)

- Confirmación durable de ganadores (Etapa 2).
- Envío en tandas sin timeout + estado real desde la planilla (Etapa 3).
- Botón "Reparar y enviar".
