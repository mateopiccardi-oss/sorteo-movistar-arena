# Entradas por ganador + envío por lotes con reanudación

**Fecha:** 2026-07-05
**Estado:** Diseño aprobado (pendiente revisión del spec por el usuario)
**Alcance:** Frontend (`index.html`) + backend (`sorteo_script.gs`). Requiere redeploy del Apps Script.

## Contexto

Esta ronda cubre tres temas aprobados:

1. **Entradas por ganador:** cada ganador puede recibir una cantidad distinta de entradas, con un valor por defecto por evento.
2. **Envío por lotes con reanudación:** el envío ya no se corta por el timeout de ~6 min del Apps Script (hoy muere a los ~25–35 mails y desincroniza la app de la planilla — caso Fito Páez). Corresponde a la "Etapa 3" anticipada en el spec de almacenamiento (2026-06-29).
3. **Envío multi-usuario garantizado:** cualquiera de los dos operadores puede enviar (o reanudar) los mails desde su compu.

**Restricción de deploy (crítica):** el Apps Script queda desplegado como **"Ejecutar como: Yo (Mateo) + Acceso: Cualquier persona"**. NO se cambia a "Usuario que accede" — rompe CORS con el frontend. Por eso los mails siempre salen desde la cuenta de Mateo con el alias `rrhh@buenosairesarena.com`, sin importar quién dispara el envío: el multi-usuario no requiere cambios de permisos, solo que el envío sea seguro de retomar desde cualquier máquina.

## Problema

### A. Cantidad de entradas única por evento

`entradasXGan` vive en el evento ([index.html:1458](../../../index.html)) y todos los ganadores reciben lo mismo. El backend multiplica `ganadores × entradasXGan` para chequear PDFs ([sorteo_script.gs:346](../../../sorteo_script.gs)) y corta los PDFs en bloques fijos ([sorteo_script.gs:379](../../../sorteo_script.gs)). No hay forma de darle 2 entradas a un ganador y 4 a otro en el mismo show.

### B. El envío se corta por timeout

`_enviarMailsLocked` ([sorteo_script.gs:316](../../../sorteo_script.gs)) envía todos los pendientes en una sola ejecución, con `Utilities.sleep(1200)` entre mails. Con ~35+ ganadores supera el límite de ~6 min de Apps Script y Google mata la ejecución a mitad de camino: la respuesta nunca llega a la app, que queda desincronizada de la planilla.

### C. Defecto latente: reasignación de PDFs al reanudar

Los PDFs se asignan por **índice dentro del lote actual de pendientes** (`pdfs.slice(i*epg, …)`). Si un envío se corta y se reintenta, los ganadores restantes vuelven a arrancar desde `pdfs[0]` → **recibirían entradas ya enviadas a otras personas**. Hay que descontar los PDFs ya usados (quedan registrados por fila en la columna "PDF enviado").

## Decisión de diseño

- **Entradas por ganador:** columna nueva **"Entradas"** en la pestaña **Ganadores** del Sheet (fuente compartida entre los dos usuarios), editable desde la app hasta que el mail se envía. Descartados: guardar solo en `localStorage` (se pierde entre compus) y repetir nombres en la matriz de Tracking (rompe las fórmulas de conteo de victorias de las columnas A/B).
- **Envío por lotes:** el backend corta limpio al acercarse a los **4 minutos** de ejecución y responde cuántos pendientes quedan; la app vuelve a llamar automáticamente hasta terminar, mostrando progreso. Descartado: continuación server-side con triggers (más complejo, sin progreso visible).
- **PDFs:** antes de asignar, el backend excluye los PDFs que ya figuran en la columna "PDF enviado" de cualquier fila del show.

## Cambios

### Backend (`sorteo_script.gs`)

#### 1. Pestaña Ganadores: columna 10 "Entradas"

- `guardarGanadores` ([sorteo_script.gs:259](../../../sorteo_script.gs)): el header pasa a incluir `"Entradas"` (col J); si la hoja ya existe y `J1` está vacío, se escribe el header una vez. Cada fila nueva escribe `g.entradas` (entero ≥ 1).
- **Filas legacy sin valor en "Entradas":** fallback al `entradasXGan` que la app manda en el request de envío (comportamiento actual), nunca a 1 fijo.

#### 2. Endpoint nuevo: `actualizarEntradas`

- `actualizarEntradas(showId, mail, entradas)`: busca la fila del show con ese mail y estado `Pendiente` y actualiza la col J. Si no encuentra fila → `{ok:false, error}` (la app avisa y revierte).
- Acción admin (fuera de `PUBLIC_ACTIONS` → exige token PIN, igual que el resto).

#### 3. `enviarMails`: por ganador, por lotes, sin reusar PDFs

`_enviarMailsLocked` ([sorteo_script.gs:316](../../../sorteo_script.gs)) cambia así:

- **Lee la cantidad por fila** (col J, fallback al parámetro `entradasXGan`) al armar la lista de pendientes.
- **PDFs usados:** arma un `Set` con los nombres de la col "PDF enviado" (col I) de **todas** las filas del show (cualquier estado) y los excluye de la lista devuelta por `buscarPDFs`.
- **Chequeo estricto:** `necesarios = suma de entradas de los pendientes`; si `pdfsDisponibles < necesarios` → no envía ninguno (regla actual, ahora con suma).
- **Asignación:** recorre los pendientes consumiendo `g.entradas` PDFs de la lista disponible, en orden alfabético (orden actual).
- **Corte por tiempo:** el `forEach` pasa a `for`; antes de cada mail, si `Date.now() - inicio > 240000` (4 min) corta el loop. El margen de ~2 min sobre el límite de Google absorbe el peor caso de un mail lento.
- **Respuesta:** agrega `restantes` = cantidad de filas del show que siguen `Pendiente` al terminar la ejecución. `{ok, enviados, enviadosMails, errores, restantes, mensaje}`.
- El lock existente por ejecución se mantiene (cada tanda toma y suelta el lock; dos usuarios no pueden pisarse).
- La secuencia por fila `Enviando → sendEmail → Enviado` no cambia: es lo que hace segura la reanudación sin duplicar mails.

#### 4. `checkPDFs`: expected desde la planilla

`checkPDFs` ([sorteo_script.gs:446](../../../sorteo_script.gs)) deja de confiar en `pendCount × entradasXGan` del frontend: lee las filas `Pendiente` del show en Ganadores y calcula `expected = suma de entradas de los pendientes` y `found = PDFs de la carpeta menos los ya usados` (mismo `Set` del punto 3, extraído a un helper compartido); `status = "ok"` si `found ≥ expected`. Mantiene `pendCount` en la respuesta para la alerta de desincronización de la app.

#### 5. `trackingGanadores`: tickets por suma

([sorteo_script.gs:750](../../../sorteo_script.gs)) `ticketsAdded` pasa de `ganadores.length × entradasXGan` a `suma de g.entradas` (fallback: `entradasXGan`). La matriz de nombres no cambia (un nombre por ganador, sin repetir).

### Frontend (`index.html`)

#### 6. Campo `entradas` en el ganador

- Al confirmar ganadores, cada objeto ganador nace con `entradas: ev.entradasXGan || 1`.
- `guardarEnSheets` ([index.html:3611](../../../index.html)) manda `entradas` por ganador; `trackingEnSheets` ([index.html:3616](../../../index.html)) manda los ganadores con `entradas` para que el backend sume.

#### 7. Stepper − / + en Ganadores confirmados

- En `renderGanConf` ([index.html:3648](../../../index.html)), cada fila no enviada muestra `− [n] +` junto al estado. Enviado → solo texto "🎟 n".
- Al tocar − / +: actualiza `g.entradas` local, `save()`, y llama (con debounce ~600 ms) a `actualizarEntradas`. Si la llamada falla → revierte el valor local y muestra toast de advertencia.
- La card de mail en `renderMails` ([index.html:3728](../../../index.html)) muestra la cantidad como texto ("🎟 2 entradas"), sin edición ahí.

#### 8. Conteos por suma

Todos los lugares que hoy hacen `ev.entradasXGan × ganadores` pasan a sumar `g.entradas ?? ev.entradasXGan ?? 1`: tickets del dashboard ([index.html:2674](../../../index.html)), conteo por evento ([index.html:1865](../../../index.html), [index.html:3367](../../../index.html), [index.html:2235](../../../index.html)).

#### 9. Envío con auto-reanudación y progreso

`_enviarMailsCore` ([index.html:3839](../../../index.html)) pasa a un loop:

```
total = pendientes; acumEnviados = 0;
do {
  res = api enviarMails
  marcar localmente los de res.enviadosMails
  acumEnviados += res.enviados
  btn.textContent = "⏳ Enviando… " + acumEnviados + "/" + total
} while (res.ok && res.restantes > 0 && iter < 30)
```

- La alerta de desincronización actual (esperaba pendientes y el backend devolvió 0 sin errores, [index.html:3847](../../../index.html)) solo aplica en la **primera** vuelta.
- Si una vuelta falla (red / `!res.ok`): se marca lo confirmado hasta ahí, toast con el error, y el botón vuelve a "Enviar via Gmail" — apretar de nuevo **reanuda** los pendientes (es el mismo flujo, no hace falta un botón aparte).
- Los errores por mail individual (`res.errores`) se acumulan entre vueltas y se muestran al final, como hoy.

### 10. Multi-usuario: verificación (sin cambios de código)

Checklist a ejecutar con el segundo usuario al cerrar la implementación:

1. Abre la app en su compu e ingresa el PIN.
2. Ajusta la cantidad de entradas de un ganador pendiente → el cambio aparece en la col J del Sheet y en la compu de Mateo al recargar.
3. Dispara un envío de prueba de punta a punta (show de prueba con 2–3 ganadores) → los mails salen desde `rrhh@buenosairesarena.com`.
4. Simulacro de reanudación: envío interrumpido en una compu se retoma desde la otra sin duplicar mails.

## Casos borde

- **Fila "Enviando" colgada** (ejecución matada justo entre `Enviando` y `Enviado`): no se reintenta automáticamente (podría duplicar el mail); queda visible en la planilla para resolución manual. El corte por tiempo a 4 min hace que este caso sea excepcional. Fuera de alcance automatizarlo.
- **Dos usuarios envían a la vez:** el lock de script serializa; el segundo recibe "Otro envío está en curso". Sin cambios.
- **PDFs con nombres repetidos en la carpeta:** el `Set` de usados los excluiría a todos. Se documenta como requisito operativo: nombres de PDF únicos por carpeta de show (ya es la práctica: `entrada_01.pdf`, `entrada_02.pdf`…).
- **Cantidad editada mientras el otro usuario envía:** `actualizarEntradas` solo toca filas `Pendiente`; si la fila ya pasó a `Enviando/Enviado`, devuelve error y la app revierte.

## Pruebas

1. **Cantidades mixtas:** show con 3 ganadores (2+2+4) → `checkPDFs` exige 8; envío adjunta 2, 2 y 4 PDFs sin superposición; col I registra los nombres correctos.
2. **Reanudación sin duplicados:** simular corte (bajar el budget a ~10 s en prueba) → segunda vuelta retoma solo `Pendiente` y **no reusa** PDFs de la col I.
3. **Persistencia multi-compu:** ajustar entradas en compu A → recargar en compu B → mismo valor; envío desde B usa las cantidades del Sheet.
4. **Legacy:** fila vieja sin col J → se envía con el `entradasXGan` del evento (comportamiento actual intacto).
5. **Progreso:** envío de >20 mails muestra el contador acumulado y termina con el total correcto en app y planilla.

## Despliegue

1. Actualizar `sorteo_script.gs` y crear **nueva implementación** manteniendo **"Ejecutar como: Yo" + "Cualquier persona"** (¡no cambiar el modo!).
2. Frontend: push a GitHub Pages (admin). No toca el formulario público ni Netlify.
3. Orden seguro: backend primero (retro-compatible: sin col J usa fallback), frontend después.

## Fuera de alcance (mejoras propuestas, no incluidas en esta ronda)

- Auditoría de quién confirmó/envió (columna "operador").
- Mail de prueba antes del envío masivo.
- Plantilla de mail editable desde el Sheet.
