# Diseño: Restyling premium de la app admin (estilo A — refinamiento)

**Fecha:** 2026-06-11
**Estado:** Aprobado (validado con mockups en el visual companion)

## Contexto

La app admin (`index.html`) tiene una identidad sólida (azul noche + cyan, Barlow/Barlow Condensed) pero terminación plana: esquinas chicas, sombras inexistentes, badges rectangulares, sin estados de foco. El usuario eligió la dirección **A — Refinamiento**: misma identidad y mismo layout, terminación premium. Se descartaron las direcciones "Obsidiana editorial" (cambiaba la identidad) y "Noche de show" (demasiado cargada para uso diario).

**Alcance: solo `index.html`** (app admin). El formulario público y `s/index.html` no se tocan. No cambia ninguna funcionalidad salvo una excepción pedida por el usuario: **se elimina la stat "Participación" del dashboard** (la métrica es relativa y no aporta).

## Tokens (variables en `:root`)

| Token | Hoy | Nuevo |
|---|---|---|
| `--r` (cards) | 8px | 12px |
| `--rsm` (botones/inputs/chips) | 5px | 9px |
| Sombra de tarjeta | — | `0 10px 26px rgba(0,0,0,0.32), inset 0 1px 0 rgba(255,255,255,0.06)` (nueva var `--shadow-card`) |
| Fondo de tarjeta | plano `rgba(22,32,48,0.9)` | `linear-gradient(160deg, rgba(26,40,64,0.95), rgba(17,26,41,0.95))` (nueva var `--grad-card`) |
| Glow primario | — | `0 6px 20px rgba(0,212,255,0.32)` (nueva var `--glow-ice`) |

Colores, fuentes y tamaños de fuente existentes **no cambian** (salvo título de página 30→32px).

## Componentes

- **`.card` y `.stat`**: fondo con gradiente, sombra en capas con luz superior (inset), borde `rgba(255,255,255,0.10)`, radio 12px. La línea cyan superior (`::before`) se mantiene.
- **Botón primario `.bp`**: gradiente `135deg, #60E8FF → #00D4FF`, radio 9px, glow `--glow-ice`; hover mantiene el lift actual.
- **Botones secundarios (`.bo`, `.bg_`, `.ba`, `.br`)**: radio 9px, fondo sutil `rgba(255,255,255,0.03)` en `.bo`.
- **Badges `.bic`**: pill (radio 99px), mantienen colores actuales.
- **Inputs `.field`**: radio 9px; focus con borde cyan + ring `0 0 0 3px rgba(0,212,255,0.10)`.
- **Topbar `.tb`**: blur 14px+, fondo `rgba(13,19,30,0.85)`; botones de nav con radio 8px (ya tienen hover/estado activo — solo se redondean).
- **Wizard `.ws`**: paso activo con círculo en gradiente cyan + glow y degradado sutil bajo el tab; paso completado igual que hoy (verde).
- **Listas de ganadores/participantes**: avatar circular con iniciales (2 letras del nombre) al inicio de cada fila; cyan para destacados/ganadores, gris para el resto. Filas con hover `rgba(255,255,255,0.03)` y radio 8px.
- **Toast**: estilo glass — fondo oscuro `rgba(13,24,32,0.95)`, borde del color del estado, sombra profunda, radio 10px.
- **Modales**: radio 14px, sombra profunda `0 24px 60px rgba(0,0,0,0.5)` (algunos ya la tienen — unificar).
- **Scrollbar**: fina y oscura (`::-webkit-scrollbar` 10px, thumb `rgba(255,255,255,0.12)` redondeado).
- **Transiciones**: las existentes (.15s) se mantienen; no se agregan animaciones de entrada nuevas.

## Cambio funcional: stat "Participación"

Verificado en el código: la stat de participación **no existe en el dashboard actual** (el mockup la inventó; los KPIs reales son Eventos / Inscriptos / Ganadores). El pedido del usuario se cumple así: **no se agrega** ninguna métrica de participación, y se elimina la función muerta `renderTasa()` (index.html:2568-2601), que calcula esa métrica para elementos `tasa-*` que ya no existen en el HTML y que ninguna parte del código llama.

## Implementación

Todo dentro de `index.html`: edición de clases en el bloque `<style>` + ajustes puntuales en los renders JS (avatares en listas, stat de participación). No se agregan dependencias ni archivos.

## Verificación

- `node --check` del bloque `<script>`.
- Revisión visual local (abrir `index.html`): dashboard, wizard completo, modal de envío de mails, historial, toasts.
- El dashboard no debe mostrar más "Participación" y las 3 stats restantes deben ocupar la fila completa.

## Fuera de alcance

- Formulario público y links cortos.
- Cambios de layout, navegación o funcionalidad (salvo la stat eliminada).
- Separar CSS/JS en archivos propios.
