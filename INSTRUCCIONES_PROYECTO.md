# Instrucciones del proyecto (pegar en Cowork → Instrucciones del proyecto)

---

## Estilo de respuesta (ahorro de tokens)
- Responde siempre en español, breve y directo. Máximo 3–5 frases salvo que pida detalle.
- No expliques el código que escribes ni repitas su contenido en el chat; solo di qué cambiaste y dónde.
- No muestres planes largos, listados de archivos ni resúmenes extensos al final.
- No releas archivos completos si solo necesitas una sección; usa búsquedas puntuales.
- No crees archivos extra (README, docs, ejemplos) salvo que los pida.
- Si algo falla, di el error en una línea y la solución propuesta, sin trazas largas.

## Contexto del proyecto
- App: IPRECON Bot (`iprecon_bot.py`) — Python, CustomTkinter + Playwright + openpyxl.
- Módulos: IMAGINE (ARL Bolívar) y GUARDIAN (Guardián de la Productividad).
- Reglas fijas: login manual (nunca guardar credenciales), solo lectura de los sitios web, no sobreescribir celdas con valor (salvo "PENDIENTE" o vacías).
- Nunca cambies la lógica de automatización ni los parsers sin que yo lo pida explícitamente.

## Ejecutables portables (Windows y Mac)
Cuando pida generar los ejecutables:
- Usar **PyInstaller** en modo ventana (sin consola): `--noconsole --onefile` en Windows (`IPRECON Bot.exe`) y `--windowed` en Mac (genera `IPRECON Bot.app`, comprimir en .zip para distribuir).
- El usuario debe poder abrirlo con doble clic, sin terminal.
- Incluir un `build.bat` (Windows) y `build_mac.sh` (Mac) que: creen venv, instalen requirements + pyinstaller, ejecuten `playwright install chromium`, y compilen.
- Empaquetar el navegador de Playwright junto al ejecutable (add-data de la carpeta de browsers o setear `PLAYWRIGHT_BROWSERS_PATH=0` antes de `playwright install` para que quede dentro del paquete) — el ejecutable debe funcionar en un PC sin Python ni internet para instalar nada.
- Importante: PyInstaller **no hace compilación cruzada**: el .exe se genera en Windows y el .app en Mac. Esta Mac genera la versión Mac; para la versión Windows, dejar listos los archivos `.spec` y `build.bat` para correrlos en un PC Windows.
- En Mac, si Gatekeeper bloquea la app sin firmar: clic derecho → Abrir, o `xattr -cr "IPRECON Bot.app"`. Documentarlo en una nota breve junto al ejecutable.
