# Entradas variables por persona (entrega contigua) — Design

**Fecha:** 2026-06-19
**Estado:** diseñado, NO implementado (Mateo lo guarda para más adelante).
**Origen:** pedido de poder dar distinta cantidad de entradas a distintos ganadores del mismo show (ej. 2 para todos, 4 para una persona), con entrega de butacas contiguas.

## Problema / objetivo

Hoy la cantidad de entradas por ganador (`entradasXGan`) es **un solo número por show**, aplicado igual a todos ([sorteo_script.gs:346](sorteo_script.gs),[:379-380](sorteo_script.gs)). Se quiere poder **sobrescribir la cantidad por persona** en un envío puntual (ej. 4 a una, 2 al resto) y que cada persona reciba **entradas contiguas** (PDFs consecutivos de la carpeta de Drive).

Las entregas ya son contiguas hoy: a cada ganador se le asigna un bloque contiguo de PDFs en orden (`pdfStart = i * entradasXGan`). Lo que falta es soportar cantidades **distintas** por persona manteniendo esa contigüidad.

## Decisiones tomadas (brainstorming)

- **Dónde se fija:** en la **pantalla de Envío** (paso 4), un campo por ganador.
- **Persistencia:** **solo al enviar** — no se guarda en la planilla. Se pasa el dato al backend en el momento del envío. (Más simple, sin migración, sin columna nueva.)

## Arquitectura

### 1. UI (pantalla de Envío, `index.html`)

Cada ganador **pendiente** muestra un input numérico con su cantidad, **pre-cargado con el default del show** (`ev.entradasXGan`, ej. 2). El usuario edita solo los que quiera. Mínimo 1 por persona.

Un **total en vivo** arriba: *"Se repartirán N entradas a M ganadores"* (suma de las cantidades), para que el usuario sepa cuántos PDFs necesita en Drive.

El estado de las cantidades editadas vive en memoria de la pantalla de envío (no en `S.ganadores`, no en localStorage). Se arma como un **mapa por mail** al momento de enviar.

### 2. Flujo de datos

Al apretar "Enviar via Gmail", el frontend arma `entradasPorMail` = `{ mail: cantidad }` solo con las personas cuya cantidad difiere del default (o con todas; equivalente). Lo manda a:
- `checkPDFs` (preflight): para validar que los PDFs alcanzan.
- `enviarMails`: para el reparto real.

Los ganadores cuyo mail **no** está en el mapa usan el default del show → comportamiento idéntico al actual.

### 3. Backend (`sorteo_script.gs`) — cambio principal

`enviarMails(showId, entradasXGan, entradasPorMail)` y `checkPDFs(showId, showNombre, entradasXGan, pendCount, entradasPorMail)` dejan de asumir "todos × N":

- **Cantidad por ganador:** `cantidad(g) = entradasPorMail[g.mail] || entradasXGan` (default del show como fallback).
- **Total necesario:** `suma de cantidad(g)` sobre los pendientes (hoy `ganadores.length * entradasXGan`). Si los PDFs en Drive son menos que la suma → no se envía ninguno (regla estricta actual, [sorteo_script.gs:347-353](sorteo_script.gs)).
- **Reparto contiguo (acumulativo):** se reemplaza `pdfStart = i * entradasXGan` por un offset acumulado: antes de cada ganador, `offset += cantidad(anterior)`; el ganador toma `pdfs.slice(offset, offset + cantidad(g))`. Así cada persona recibe PDFs **consecutivos** aunque las cantidades difieran, en el orden de las filas de Ganadores.

### 4. Validación y errores

- Mínimo 1 entrada por persona (la UI no permite 0 ni vacío).
- El total no puede superar los PDFs disponibles → lo frena el preflight (`checkPDFs`), mismo flujo de hoy (status `insufficient`).
- `entradasPorMail` vacío o ausente, o un mail no listado → se usa el default del show. **Compatible hacia atrás**: sin mapa, comportamiento byte-idéntico al actual.

### 5. Despliegue

Es el **primer cambio al backend** de esta tanda de trabajo. Requiere **re-desplegar el Apps Script**, manteniendo el modo *"Ejecutar como: Yo + Cualquier persona"* (NO "Usuario que accede" — rompe el CORS con el frontend de Netlify). Frontend (`index.html`) y backend (`sorteo_script.gs`) se suben juntos.

## Testing

Sin framework de tests. Verificación:
- `node --check` del `<script>` (frontend) y revisión estática del `.gs`.
- Manual en navegador: un show con varios pendientes; poner 2 a la mayoría y 4 a uno; verificar el total en vivo; preflight con PDFs justos/insuficientes; y (con PDFs reales ordenados por butaca) confirmar que cada persona recibe el bloque contiguo correcto.

## Fuera de alcance

- Guardar la cantidad por persona en la planilla (se descartó: "solo al enviar").
- Cantidad 0 / excluir a alguien del envío.
- Reordenar ganadores para controlar qué butacas recibe cada uno (se usa el orden de las filas de Ganadores).
