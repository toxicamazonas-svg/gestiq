# Gestiq — Sistema de diseño "Cristal Disciplinado"

_v1 · 02-jul-2026 · Aplica a `gestiq_ui.html` (app) y `docs/cuenta.html` (web de cuenta)._
Referencias: Arc, macOS moderno. Carácter: glass translúcido ejecutado con sistema — un solo verde, bordes hairline, aire consistente, motion sobrio.

## 1 · Tokens de color

Semánticos, definidos en `:root` (oscuro, por defecto) y `body.claro`.

| Token | Oscuro | Claro | Uso |
|---|---|---|---|
| `--bg` | `#101018` | `#F2F3F6` | fondo de ventana |
| `--card` | `rgba(255,255,255,.055)` | `rgba(255,255,255,.66)` | superficie glass (tarjetas, nav, header) |
| `--card-op` | `#1A1A24` | `#FFFFFF` | superficie opaca (options, popovers) |
| `--inner` | `rgba(255,255,255,.045)` | `rgba(16,18,32,.045)` | campos y zonas dentro de tarjeta |
| `--text` | `#F2F3F8` | `#16181F` | texto principal |
| `--text2` | `#A6ABC2` | `#565C6E` | secundario (labels, ayuda) |
| `--text3` | `#6E7389` | `#868CA0` | terciario (solo ≥12px y no esencial) |
| `--brd` | `rgba(255,255,255,.10)` | `rgba(16,18,32,.10)` | borde hairline |
| `--brd2` | `rgba(255,255,255,.16)` | `rgba(16,18,32,.18)` | borde interactivo/hover |
| `--brand` | `#2BD983` | `#0FA65C` | acento de marca |
| `--brand-deep` | `#0FA65C` | `#0B7A45` | extremo del gradiente de marca |
| `--on-brand` | `#052616` | `#FFFFFF` | texto sobre marca (tinta, no blanco, en oscuro) |
| `--ok` / `--warn` / `--err` | `#2BD983` / `#FFB454` / `#FF6369` | `#0E9F5D` / `#B26A0E` / `#D33B3B` | semánticos |
| `--tint-ok/-warn/-err` | rgba de cada uno al `.12` | al `.10` | fondos de chips/badges |
| `--log-bg` | `rgba(8,8,14,.6)` | `#14171E` | consola (siempre oscura) |

Reglas: un solo par de verdes (adiós `#00B050/#00C75B/#3DDC84` mezclados); un solo rojo por modo; gradiente de marca = `linear-gradient(160deg, var(--brand), var(--brand-deep))` y nada más. Fondo ambiental: los 3 radial-gradient existentes, bajados a verde `.10` máx.

## 2 · Tipografía

Stack de sistema (app nativa, sin red): `-apple-system, "SF Pro Display", "Segoe UI Variable", "Segoe UI", Inter, sans-serif`. Mono: `"SF Mono", "Cascadia Mono", Menlo, Consolas, monospace`.

| Rol | Tamaño | Peso | Extra |
|---|---|---|---|
| Título de página (h1) | 17px | 700 | letter-spacing -.02em |
| Título de tarjeta (h2) | 13px | 650 | -.01em |
| Cuerpo / controles | 13px | 500 | |
| Secundario | 12px | 500 | color `--text2` |
| Caption / secciones nav | 10.5px | 650 | letter-spacing .1em, MAYÚSC., `--text3` |
| Log / datos | 11.5px mono | 400 | line-height 1.75 |

Pesos permitidos: 400/500/650/700 (hoy todo es 700–800 → nada resalta). `font-variant-numeric: tabular-nums` en versión, progreso, conteos y log.

## 3 · Espaciado, radios, sombras, grid

- **Espaciado** base 4: `4 · 8 · 12 · 16 · 20 · 24 · 32`. Gap global de layout 12. Padding tarjeta 16×18. Nada de 7/9/11/13.
- **Radios**: `--r-s:9` (iconos-botón, celdas), `--r-m:14` (campos, log, zonas internas, chip sesión), `--r-l:22` (tarjetas, nav, modales), `99` (header, botones píldora, chips, avatar).
- **Sombras**: superficies NO llevan drop-shadow (el borde separa; decisión existente que se mantiene). Solo flotantes: modal/toast `0 30px 80px rgba(0,0,0,.5)`, glider y CTA `0 5px 16px rgba(brand,.30)`. Highlight interno `inset 0 1px 0 rgba(255,255,255,.08)` en todo glass.
- **Blur**: `blur(24px) saturate(160%)` en header/nav/tarjetas; `blur(14px)` en backdrops. Máx. 2 capas de blur visibles a la vez.
- **Grid app**: header píldora 48px · sidebar 204px · main fluido, gap 12, padding ventana 14/16. Mín. 960×660 soportado sin recortes.
- **Grid cuenta.html**: mobile-first, contenedor 420px centrado, touch targets ≥44px.

## 4 · Iconografía

SVG de línea inline (sprite `<symbol>` en el HTML): stroke `1.8`, `currentColor`, linecap/linejoin round, caja 24, render 15–17px. Set: lupa (IMAGINE), escudo (GUARDIÁN), documento, play, stop, descarga, engranaje, sol/luna, exportar, refrescar, check, alerta, x, salir, usuario, reloj, ojo.
**Prohibido**: emojis como iconos (📄▶⏹⚙🌙ℹ️ fuera) y los PNG de nav — los módulos usan lupa/escudo de línea que heredan color de estado. _(Suposición aprobable: se retiran los logos PNG de IMAGINE/GUARDIÁN de la nav; siguen siendo los nombres.)_

## 5 · Componentes base

- **Botón primario**: píldora 36px, gradiente de marca, texto `--on-brand` 650, sombra tintada. Hover: brillo +6% y -1px translateY. Active: scale .98. Disabled: fondo `--inner`, texto `--text3`, sin sombra ni gradiente (ya no "verde apagado").
- **Botón secundario**: `--inner` + borde `--brd`, texto `--text`. **Peligro**: outline rojo, se rellena `--tint-err` al hover (sólido solo mientras corre una tarea). **Ghost/enlace**: sin caja, subrayado al hover.
- **Botón-icono**: 32×32 (`--r-s` o 99), `--inner` + borde. Sin rotate en hover.
- **Campos** (input/select): 36px, `--inner`, borde `--brd`, radio 14. Focus: borde brand + ring `0 0 0 3px rgba(brand,.22)`. Invalid: borde `--err` + mensaje 12px debajo. Select con flecha SVG propia.
- **Chips de hoja**: píldora 28px; base `--inner`+borde; seleccionada = gradiente marca + texto `--on-brand`; "Todas" borde dashed. Contenedor con marco `--inner` (se mantiene) y fade inferior si scrollea.
- **Chips de resumen**: píldora 24px, tinte semántico de fondo + texto del color pleno.
- **Barra de progreso**: 7px, riel `--inner`+borde, relleno gradiente marca + glow `rgba(brand,.45)`.
- **Log**: `--log-bg`, radio 14, mono 11.5px, timestamps `--text3`, seleccionable. Prefijos ✓/⚠/✕ coloreados, texto neutro.
- **Toasts**: **abajo-derecha** (hoy tapan el header), radio 14, borde-l 3px semántico, autocierre con barrita, máx 3 visibles.
- **Modales**: radio 22, backdrop `rgba(10,10,18,.45)` + blur 14, entrada spring única.
- **Segmentado** (tema/módulo): riel `--inner` píldora; opción activa gradiente marca.
- **Badge de estado de paso**: "● Cargado / Pendiente" 11px en el encabezado de cada tarjeta.
- **Empty states**: icono línea 20px `--text3` + frase corta + hint de acción (zona hojas, log vacío).
- **Skeleton**: bloques `--inner` con shimmer sutil 1.4s al abrir archivo (zona archivo + hojas).

## 6 · Estados y accesibilidad

- Todo interactivo tiene los 6 estados: normal / hover / active / **focus-visible** (`outline:2px solid var(--brand); outline-offset:2px`) / disabled / loading (spinner 14px en botón, texto "Verificando…" etc.).
- Contraste AA verificado por par en ambos temas (`--text3` solo ≥12px). Texto mínimo 11px (solo captions).
- Hit-area ≥28px en app escritorio, ≥44px en cuenta.html móvil.
- `user-select`: permitido en correo, log, valores y mensajes de error (hoy el `none` global lo impide).
- Los pasos 2·3 se atenúan (`opacity .45` + `pointer-events` limitados) hasta cargar archivo: el flujo 1→2→3 se lee solo.

## 7 · Motion

Duraciones 120 (hover) / 200 (estado) / 320ms (entrada). Easing `--suave: cubic-bezier(.22,1,.36,1)`; spring solo en glider de nav y entrada de modal. La cascada de entrada de tarjetas corre **solo al arrancar** (no se re-anima al cambiar de módulo — defecto actual). Cambio de tema: view-transition circular existente se conserva. Sin `rotate()` en hovers. `prefers-reduced-motion`: se respeta desactivando transforms.
