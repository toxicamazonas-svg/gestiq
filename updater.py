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
import hashlib
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

# Estado compartido entre el hilo de fondo, la UI (progreso) y el cierre.
# fase: nada | descargando | verificando | preparando | listo | error
_ESTADO = {"listo": False, "tipo": None, "tmp": None, "tag": None,
           "fase": "nada", "pct": None, "mb": 0.0, "mb_total": None,
           "error": "", "aplicado": False}

# Candado: SOLO una descarga a la vez. Sin esto, el clic en "Buscar
# actualizaciones" lanzaba una segunda descarga al MISMO directorio mientras
# la del arranque seguía corriendo; cada una hace rmtree del trabajo de la
# otra y en Mac/POSIX ambas mueren (os.replace sobre ruta borrada) → el
# update "nunca quedaba listo" y al cerrar no se instalaba nada.
_DL_LOCK = threading.Lock()

# Aviso de "descarga lista" registrado una sola vez (main lo pasa al arrancar);
# lo reutiliza cualquier descarga, incluida la que dispara la comprobación manual.
_ON_LISTO = {"cb": None}


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
    assets, digests = {}, {}
    for a in (data.get("assets") or []):
        assets[a.get("name")] = a.get("browser_download_url")
        # GitHub publica el sha256 del asset como "digest": "sha256:<hex>"
        digests[a.get("name")] = str(a.get("digest") or "")
    notas = (data.get("body") or "").strip()      # changelog de la release
    return tag, assets, digests, notas


def _sha256_ok(path, digest):
    """True si el archivo coincide con el digest ('sha256:<hex>') o si el
    release no trae digest (no se puede verificar → no bloquear el update)."""
    d = (digest or "").strip().lower()
    if not d.startswith("sha256:"):
        return True
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest() == d.split(":", 1)[1].strip()


def _descargar(url, dst):
    """Descarga por bloques publicando el progreso en _ESTADO (mb/mb_total/pct)
    para que la barra del modal muestre el avance real."""
    tmp = dst + ".part"
    with urllib.request.urlopen(
            urllib.request.Request(url, headers={"User-Agent": "Gestiq-Updater"}),
            timeout=120, context=_ctx()) as r, open(tmp, "wb") as f:
        try:    total = int(r.headers.get("Content-Length") or 0) or None
        except Exception: total = None
        _ESTADO.update(mb_total=round(total / 1048576, 1) if total else None)
        leidos = 0
        for bloque in iter(lambda: r.read(1 << 18), b""):      # 256 KB
            f.write(bloque)
            leidos += len(bloque)
            _ESTADO.update(mb=round(leidos / 1048576, 1),
                           pct=min(99, leidos * 100 // total) if total else None)
    os.replace(tmp, dst)


# ── Descarga del update (hilo de fondo) ───────────────────────────────────────
def _buscar_y_descargar(on_listo=None):
    if on_listo is None:
        on_listo = _ON_LISTO.get("cb")
    if _ESTADO.get("listo"):
        return                      # ya hay un update descargado y listo
    if _translocada():
        return                      # volumen de solo lectura: imposible actualizar
    if not _DL_LOCK.acquire(blocking=False):
        return                      # ya hay una descarga en curso: no pisarla
    try:
        _descargar_con_candado(on_listo)
    finally:
        _DL_LOCK.release()


def _descargar_con_candado(on_listo):
    info = _info()
    if not info:
        return
    plat, _objetivo, base = info
    try:
        tag, assets, digests, _notas = _release_latest()
    except Exception:
        return
    if not tag or not _es_mas_nueva(tag, VERSION):
        return

    if plat == "win":
        nombre = ASSET_WIN_LIGERO if assets.get(ASSET_WIN_LIGERO) else ASSET_WIN_FULL
    else:
        nombre = ASSET_MAC
    url = assets.get(nombre)
    if not url:
        return

    work = os.path.join(base, ".gestiq_update")
    try:
        _ESTADO.update(fase="descargando", pct=None, mb=0.0, mb_total=None,
                       error="", tag=tag)
        shutil.rmtree(work, ignore_errors=True)
        os.makedirs(work, exist_ok=True)
        zpath = os.path.join(work, "paquete.zip")
        _descargar(url, zpath)
        _ESTADO.update(fase="verificando")
        if not _sha256_ok(zpath, digests.get(nombre)):
            raise RuntimeError("SHA256 del update no coincide — descarga descartada")
        _ESTADO.update(fase="preparando")

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

        _ESTADO.update(listo=True, tipo=plat, tmp=nuevo, tag=tag,
                       fase="listo", pct=100)
        if on_listo:
            try:
                on_listo(tag)
            except Exception:
                pass
    except Exception as e:
        _ESTADO.update(fase="error", error=str(e)[:80], pct=None)
        shutil.rmtree(work, ignore_errors=True)


def iniciar_en_segundo_plano(on_listo=None):
    """Lanza la comprobación de actualizaciones sin bloquear el arranque."""
    if on_listo:
        _ON_LISTO["cb"] = on_listo
    threading.Thread(target=_buscar_y_descargar, daemon=True).start()


def _translocada():
    """True si macOS está ejecutando la app desde App Translocation (la copia
    aún tiene la cuarentena y corre desde un volumen temporal de SOLO lectura):
    ahí el updater no puede ni descargar. Se cura con `xattr -cr Gestiq.app`."""
    return sys.platform == "darwin" and "/AppTranslocation/" in (sys.executable or "")


def comprobar_ahora():
    """Comprobación manual (botón de la UI). Devuelve un dict para mostrar:
    {hay, tag, actual, notas, listo, dev, transloc} o {error}. Si hay una
    versión nueva y aún no se había descargado, arranca la descarga."""
    try:
        tag, _assets, _digests, notas = _release_latest()
    except Exception as e:
        return {"error": f"No se pudo consultar las versiones ({str(e)[:60]})"}
    hay = bool(tag) and _es_mas_nueva(tag, VERSION)
    r = {"hay": hay, "tag": tag, "actual": VERSION,
         "notas": notas[:4000], "listo": hay_update_listo(),
         "dev": not getattr(sys, "frozen", False),
         "transloc": _translocada()}
    if hay and not _ESTADO.get("listo") and getattr(sys, "frozen", False) \
           and not _translocada():
        iniciar_en_segundo_plano()                 # empieza a bajarla ya
    return r


def hay_update_listo():
    return bool(_ESTADO.get("listo"))


def estado_update():
    """Progreso de la descarga para la UI del modal (se sondea):
    {fase, pct (0-100 o None), mb, mb_total, tag, error, listo}."""
    e = _ESTADO
    return {"fase": "listo" if e.get("listo") else e.get("fase", "nada"),
            "pct": 100 if e.get("listo") else e.get("pct"),
            "mb": e.get("mb"), "mb_total": e.get("mb_total"),
            "tag": e.get("tag"), "error": e.get("error", ""),
            "listo": bool(e.get("listo"))}


# ── Aplicación del update (al cerrar) ─────────────────────────────────────────
def aplicar_y_reiniciar():
    """Si hay un update descargado, lanza el ayudante que reemplaza y relanza.

    Devuelve True si lanzó el proceso de actualización."""
    if not _ESTADO.get("listo"):
        return False
    if _ESTADO.get("aplicado"):
        return True                 # ya se lanzó el ayudante (p.ej. "Actualizar
                                    # ahora" + evento closing): no lanzar DOS
    info = _info()
    if not info:
        return False
    plat, objetivo, base = info
    nuevo = _ESTADO.get("tmp")
    if not nuevo or not os.path.exists(nuevo):
        return False
    pid = os.getpid()
    work = os.path.join(base, ".gestiq_update")
    # Los ayudantes viven en la carpeta TEMPORAL del sistema (no en work):
    # así pueden borrar work al terminar sin serrucharse la rama (un script
    # no puede eliminar con seguridad la carpeta desde la que se ejecuta).
    try:
        if plat == "win":
            bat = os.path.join(tempfile.gettempdir(), f"gestiq_aplicar_{pid}.bat")
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
                    f'rmdir /S /Q "{work}" >NUL 2>&1\r\n'
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
            sh = os.path.join(tempfile.gettempdir(), f"gestiq_aplicar_{pid}.sh")
            with open(sh, "w", encoding="utf-8") as f:
                f.write(
                    "#!/bin/bash\n"
                    f"while kill -0 {pid} 2>/dev/null; do sleep 1; done\n"
                    "sleep 1\n"
                    # Respaldo y reemplazo con vuelta atrás: si el mv falla,
                    # se restaura la app anterior (nunca dejar al usuario sin app).
                    f'rm -rf "{objetivo}.old"\n'
                    f'mv "{objetivo}" "{objetivo}.old"\n'
                    f'if mv "{nuevo}" "{objetivo}"; then\n'
                    f'  xattr -cr "{objetivo}"\n'
                    f'  rm -rf "{objetivo}.old"\n'
                    # instalado: fuera la carpeta de trabajo (antes quedaba
                    # .gestiq_update residual junto a la app)
                    f'  rm -rf "{work}"\n'
                    "else\n"
                    f'  mv "{objetivo}.old" "{objetivo}"\n'
                    "fi\n"
                    f'open "{objetivo}"\n'
                    'rm -f "$0"\n'
                )
            os.chmod(sh, 0o755)
            subprocess.Popen(["/bin/bash", sh], start_new_session=True,
                             close_fds=True)
        _ESTADO["aplicado"] = True
        return True
    except Exception:
        return False
