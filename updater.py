# -*- coding: utf-8 -*-
"""Auto-actualización de Gestiq desde GitHub Releases.

Flujo (silencioso, sin que el usuario haga nada):
  1. Al arrancar, en un hilo de fondo, consulta el último release del repo.
  2. Si hay una versión mayor, descarga el asset adecuado a una carpeta
     temporal junto al ejecutable.
  3. Al cerrar la app, si la descarga quedó lista, lanza un ayudante que
     espera a que el proceso termine, reemplaza los archivos y relanza Gestiq.

Windows: reemplaza solo Gestiq.exe (conserva ms-playwright) vía un .bat.
Mac:     reemplaza Gestiq.app completo y ejecuta `xattr -cr` (quita la
         cuarentena de Gatekeeper) vía un .sh.

Solo actúa cuando la app está compilada (sys.frozen). En desarrollo no hace nada.
Todo está envuelto en try/except: si algo falla, la app sigue funcionando.
"""

import os
import re
import sys
import json
import ssl
import shutil
import zipfile
import tempfile
import threading
import subprocess
import urllib.request

from version import VERSION

REPO = "toxicamazonas-svg/gestiq"
API_LATEST = f"https://api.github.com/repos/{REPO}/releases/latest"

ASSET_WIN_LIGERO = "Gestiq-Update-Windows.zip"   # solo el exe (update ligero)
ASSET_WIN_FULL   = "Gestiq-Windows.zip"          # zip completo (respaldo)
ASSET_MAC        = "Gestiq-Mac.zip"

# Estado compartido entre el hilo de fondo y el cierre de la app.
_ESTADO = {"listo": False, "tipo": None, "tmp": None, "tag": None}


# ── Utilidades de versión ─────────────────────────────────────────────────────
def _tupla(v):
    nums = re.findall(r"\d+", v or "")[:3]
    nums += ["0"] * (3 - len(nums))
    return tuple(int(n) for n in nums)


def _es_mas_nueva(remota, actual):
    return _tupla(remota) > _tupla(actual)


# ── Rutas del ejecutable / bundle ─────────────────────────────────────────────
def _info():
    """Devuelve (plataforma, ruta_objetivo, carpeta_base) o None si no aplica."""
    if not getattr(sys, "frozen", False):
        return None
    exe = sys.executable
    if sys.platform.startswith("win"):
        return ("win", exe, os.path.dirname(exe))
    if sys.platform == "darwin":
        # .../Gestiq.app/Contents/MacOS/Gestiq  →  .../Gestiq.app
        app = os.path.dirname(os.path.dirname(os.path.dirname(exe)))
        if app.endswith(".app"):
            return ("mac", app, os.path.dirname(app))
    return None


# ── Red ───────────────────────────────────────────────────────────────────────
def _ctx():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def _get(url, binario=False):
    req = urllib.request.Request(url, headers={
        "User-Agent": "Gestiq-Updater",
        "Accept": "application/octet-stream" if binario
                  else "application/vnd.github+json",
    })
    with urllib.request.urlopen(req, timeout=30, context=_ctx()) as r:
        return r.read()


def _release_latest():
    data = json.loads(_get(API_LATEST).decode("utf-8", "replace"))
    tag = (data.get("tag_name") or "").strip()
    assets = {a.get("name"): a.get("browser_download_url")
              for a in (data.get("assets") or [])}
    return tag, assets


def _descargar(url, dst):
    tmp = dst + ".part"
    with urllib.request.urlopen(
            urllib.request.Request(url, headers={"User-Agent": "Gestiq-Updater"}),
            timeout=120, context=_ctx()) as r, open(tmp, "wb") as f:
        shutil.copyfileobj(r, f)
    os.replace(tmp, dst)


# ── Descarga del update (hilo de fondo) ───────────────────────────────────────
def _buscar_y_descargar(on_listo=None):
    info = _info()
    if not info:
        return
    plat, _objetivo, base = info
    try:
        tag, assets = _release_latest()
    except Exception:
        return
    if not tag or not _es_mas_nueva(tag, VERSION):
        return

    if plat == "win":
        url = assets.get(ASSET_WIN_LIGERO) or assets.get(ASSET_WIN_FULL)
    else:
        url = assets.get(ASSET_MAC)
    if not url:
        return

    work = os.path.join(base, ".gestiq_update")
    try:
        shutil.rmtree(work, ignore_errors=True)
        os.makedirs(work, exist_ok=True)
        zpath = os.path.join(work, "paquete.zip")
        _descargar(url, zpath)

        if plat == "win":
            with zipfile.ZipFile(zpath) as z:
                nombre = next((n for n in z.namelist()
                               if n.lower().endswith("gestiq.exe")), None)
                if not nombre:
                    return
                z.extract(nombre, work)
                nuevo = os.path.join(work, "Gestiq.exe")
                os.replace(os.path.join(work, nombre), nuevo)
        else:
            # ditto preserva permisos y el bit ejecutable del binario interno.
            subprocess.run(["ditto", "-x", "-k", zpath, work],
                           check=True, capture_output=True)
            nuevo = None
            for raiz, dirs, _ in os.walk(work):
                for d in dirs:
                    if d.endswith(".app"):
                        nuevo = os.path.join(raiz, d)
                        break
                if nuevo:
                    break
            if not nuevo:
                return

        try:
            os.remove(zpath)
        except OSError:
            pass

        _ESTADO.update(listo=True, tipo=plat, tmp=nuevo, tag=tag)
        if on_listo:
            try:
                on_listo(tag)
            except Exception:
                pass
    except Exception:
        shutil.rmtree(work, ignore_errors=True)


def iniciar_en_segundo_plano(on_listo=None):
    """Lanza la comprobación de actualizaciones sin bloquear el arranque."""
    threading.Thread(target=_buscar_y_descargar, args=(on_listo,),
                     daemon=True).start()


def hay_update_listo():
    return bool(_ESTADO.get("listo"))


# ── Aplicación del update (al cerrar) ─────────────────────────────────────────
def aplicar_y_reiniciar():
    """Si hay un update descargado, lanza el ayudante que reemplaza y relanza.

    Devuelve True si lanzó el proceso de actualización."""
    if not _ESTADO.get("listo"):
        return False
    info = _info()
    if not info:
        return False
    plat, objetivo, base = info
    nuevo = _ESTADO.get("tmp")
    if not nuevo or not os.path.exists(nuevo):
        return False
    pid = os.getpid()
    work = os.path.join(base, ".gestiq_update")
    try:
        if plat == "win":
            bat = os.path.join(work, "aplicar.bat")
            with open(bat, "w", encoding="ascii", errors="ignore") as f:
                f.write(
                    "@echo off\r\n"
                    "rem 1) esperar a que se cierren TODOS los procesos de la app\r\n"
                    ":waitproc\r\n"
                    'tasklist /FI "IMAGENAME eq Gestiq.exe" 2>NUL | find /I "Gestiq.exe" >NUL\r\n'
                    "if not errorlevel 1 ( ping -n 2 127.0.0.1 >NUL & goto waitproc )\r\n"
                    "rem 2) reintentar el reemplazo hasta que el .exe se libere\r\n"
                    "set /a intentos=0\r\n"
                    ":trymove\r\n"
                    f'move /Y "{nuevo}" "{objetivo}" >NUL 2>&1\r\n'
                    "if not errorlevel 1 goto relanzar\r\n"
                    "set /a intentos+=1\r\n"
                    "ping -n 2 127.0.0.1 >NUL\r\n"
                    "if %intentos% lss 60 goto trymove\r\n"
                    "goto limpiar\r\n"
                    ":relanzar\r\n"
                    f'start "" "{objetivo}"\r\n'
                    ":limpiar\r\n"
                    '(goto) 2>nul & del "%~f0"\r\n'
                )
            DETACHED = 0x00000008
            NEWGROUP = 0x00000200
            NOWINDOW = 0x08000000
            subprocess.Popen(["cmd", "/c", bat],
                             creationflags=DETACHED | NEWGROUP | NOWINDOW,
                             close_fds=True, cwd=base)
        else:
            sh = os.path.join(work, "aplicar.sh")
            with open(sh, "w", encoding="utf-8") as f:
                f.write(
                    "#!/bin/bash\n"
                    f"while kill -0 {pid} 2>/dev/null; do sleep 1; done\n"
                    "sleep 1\n"
                    f'rm -rf "{objetivo}"\n'
                    f'mv "{nuevo}" "{objetivo}"\n'
                    f'xattr -cr "{objetivo}"\n'
                    f'open "{objetivo}"\n'
                    'rm -f "$0"\n'
                )
            os.chmod(sh, 0o755)
            subprocess.Popen(["/bin/bash", sh], start_new_session=True,
                             close_fds=True)
        return True
    except Exception:
        return False
