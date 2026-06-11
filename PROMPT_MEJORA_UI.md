# Prompt: Rediseño visual de IPRECON Bot

Copia y pega esto en Cowork:

---

Rediseña por completo la interfaz gráfica de `iprecon_bot.py` (CustomTkinter) para que se vea profesional, moderna y minimalista, pensada para un usuario común sin conocimientos técnicos. **No toques nada de la lógica de automatización** (Playwright, parsers, escritura en Excel): solo la capa visual.

## Estilo visual
- Tema oscuro elegante por defecto, con switch claro/oscuro arriba a la derecha.
- Paleta sobria de máximo 4 colores: un fondo neutro (#1E1E2E o similar), un color de acento corporativo (verde Bolívar #00B050 o azul), gris para texto secundario, y rojo solo para errores. Nada de colores chillones.
- Tipografía Segoe UI / SF Pro según el sistema, tamaños consistentes: títulos 18–20, texto 13, secundario 11.
- Bordes redondeados (corner_radius 10–12), espaciado generoso (padding 16–24), sin elementos amontonados.
- Iconos simples (emoji o caracteres unicode) en botones y pestañas, sin sobrecargar.

## Estructura
- Encabezado con nombre "IPRECON Bot", logo/inicial en un círculo de acento y versión.
- Sidebar izquierdo o pestañas limpias para IMAGINE y GUARDIAN, con indicador claro de la pestaña activa.
- Cada módulo organizado en tarjetas (frames con fondo levemente distinto): 1) Selección de archivo Excel, 2) Configuración, 3) Ejecución y progreso.
- Selector de archivo con zona clara que muestre el nombre del archivo elegido y un botón "Cambiar".
- Botón principal de acción grande y destacado (color de acento); botón "Detener" en gris/rojo solo visible mientras corre.

## Feedback al usuario
- Barra de progreso moderna (CTkProgressBar) con "X de Y casos" y porcentaje.
- Área de log con fondo oscuro tipo consola pero estilizada: mensajes con prefijos de color (✓ verde, ✗ rojo, ⚠ naranja), fuente monoespaciada pequeña, auto-scroll.
- Estados claros: "Esperando login manual…", "Procesando caso 12345…", "Completado: 40 ✓, 2 ✗".
- Mensajes de error en lenguaje simple, nunca trazas técnicas crudas (esas van solo al log).

## Usabilidad
- Ventana con tamaño mínimo razonable (p. ej. 900x650) y redimensionable sin que se rompa el layout (usar grid con weights).
- Deshabilitar controles mientras el bot corre, para evitar clics accidentales.
- Tooltips o textos de ayuda breves bajo cada campo.
- Confirmación al cerrar si hay un proceso en curso.

Al terminar, verifica que la app abre sin errores y que ambos módulos siguen funcionando igual que antes. Responde de forma breve, solo con el resumen de cambios.
