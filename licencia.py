# -*- coding: utf-8 -*-
"""
licencia.py — Sistema de licencias de Gestiq (Supabase, plan gratis).

Diseño:
- Login email + contraseña contra Supabase Auth (JWT + refresh token).
- El refresh token se guarda en el llavero del sistema (keyring) → el usuario
  inicia sesión UNA sola vez por equipo.
- La licencia se valida SIEMPRE en el servidor (función SQL check_license):
  estado, fecha de vencimiento (hora del servidor, no del PC) y machine_id.
- Sin internet o sin respuesta del servidor = sin acceso. No hay modo offline.

Proyecto Supabase ya configurado abajo (URL + anon key). Sin licencia
válida la app queda bloqueada — no existe modo desarrollo.
"""

import sys, os, re, json, uuid, base64, hashlib, platform, subprocess, ssl
import urllib.request
import urllib.error

# ── Configuración (Settings → API del proyecto Supabase) ────────────────────
SUPABASE_URL      = "https://xdydorreyvkenbifefus.supabase.co"
SUPABASE_ANON_KEY = "sb_publishable_podBtj6El-uznTp5gCzAYg_Tss_j1Xj"

TIMEOUT = 12          # segundos por petición
_SERVICIO = "gestiq-licencia"


class LicenciaError(Exception):
    """Error mostrable al usuario (red, credenciales, licencia)."""


def configurado():
    return SUPABASE_URL.startswith("https://") and "PEGAR" not in SUPABASE_ANON_KEY


# ── Identificador estable del equipo ─────────────────────────────────────────
def machine_id():
    raw = ""
    try:
        if sys.platform == "darwin":
            out = subprocess.run(["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"],
                                 capture_output=True, text=True, timeout=5).stdout
            m = re.search(r'IOPlatformUUID"\s*=\s*"([^"]+)"', out)
            if m: raw = m.group(1)
        elif sys.platform.startswith("win"):
            import winreg
            k = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                               r"SOFTWARE\Microsoft\Cryptography", 0,
                               winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
            raw = str(winreg.QueryValueEx(k, "MachineGuid")[0])
        else:
            with open("/etc/machine-id") as f:
                raw = f.read().strip()
    except Exception:
        pass
    if not raw:
        raw = f"{platform.node()}-{uuid.getnode()}"
    return hashlib.sha256(f"gestiq|{raw}".encode()).hexdigest()[:32]


# ── HTTP ──────────────────────────────────────────────────────────────────────
_SSL_CTX = None

def _ssl_ctx():
    """Contexto SSL con certificados; en Mac el Python oficial no trae los del
    sistema, así que se cargan los de certifi si están disponibles."""
    global _SSL_CTX
    if _SSL_CTX is None:
        _SSL_CTX = ssl.create_default_context()
        try:
            import certifi
            _SSL_CTX.load_verify_locations(certifi.where())
        except Exception:
            pass
    return _SSL_CTX


def _post(path, body, token=None):
    req = urllib.request.Request(
        SUPABASE_URL.rstrip("/") + path,
        data=json.dumps(body).encode(),
        headers={"apikey": SUPABASE_ANON_KEY,
                 "Authorization": f"Bearer {token or SUPABASE_ANON_KEY}",
                 "Content-Type": "application/json"},
        method="POST")
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT, context=_ssl_ctx()) as r:
            return r.status, json.loads(r.read().decode() or "null")
    except urllib.error.HTTPError as e:
        try:    return e.code, json.loads(e.read().decode() or "null")
        except Exception: return e.code, None
    except Exception as e:
        detalle = str(getattr(e, "reason", e))[:90]
        raise LicenciaError("Sin conexión con el servidor de licencias. "
                            "Gestiq necesita internet para funcionar.\n"
                            f"({detalle})") from e


# ── Guardado del refresh token (llavero del sistema) ─────────────────────────
def _kr():
    try:
        import keyring
        return keyring
    except Exception:
        return None

def _fallback_path():
    return os.path.join(os.path.expanduser("~"), ".gestiq_sesion")

def _guardar(clave, valor):
    kr = _kr()
    if kr:
        try: kr.set_password(_SERVICIO, clave, valor); return
        except Exception: pass
    try:
        data = {}
        if os.path.exists(_fallback_path()):
            data = json.loads(base64.b64decode(open(_fallback_path(), "rb").read()))
        data[clave] = valor
        with open(_fallback_path(), "wb") as f:
            f.write(base64.b64encode(json.dumps(data).encode()))
    except Exception:
        pass

def _leer(clave):
    kr = _kr()
    if kr:
        try:
            v = kr.get_password(_SERVICIO, clave)
            if v: return v
        except Exception:
            pass
    try:
        data = json.loads(base64.b64decode(open(_fallback_path(), "rb").read()))
        return data.get(clave)
    except Exception:
        return None

def _borrar(clave):
    kr = _kr()
    if kr:
        try: kr.delete_password(_SERVICIO, clave)
        except Exception: pass
    try:
        p = _fallback_path()
        if os.path.exists(p):
            data = json.loads(base64.b64decode(open(p, "rb").read()))
            data.pop(clave, None)
            with open(p, "wb") as f:
                f.write(base64.b64encode(json.dumps(data).encode()))
    except Exception:
        pass


# ── Sesión ────────────────────────────────────────────────────────────────────
def _sesion_desde(j):
    if not j or "access_token" not in j:
        raise LicenciaError("Respuesta de sesión no válida.")
    s = {"access_token": j["access_token"],
         "refresh_token": j.get("refresh_token", ""),
         "email": (j.get("user") or {}).get("email", "")}
    if s["refresh_token"]:
        _guardar("refresh", s["refresh_token"])
    if s["email"]:
        _guardar("email", s["email"])
    return s


def login(email, password):
    """Inicia sesión. Devuelve la sesión o lanza LicenciaError."""
    st, j = _post("/auth/v1/token?grant_type=password",
                  {"email": email.strip(), "password": password})
    if st == 200:
        return _sesion_desde(j)
    msg = (j or {}).get("error_description") or (j or {}).get("msg") or ""
    if "Invalid login" in msg or st in (400, 401):
        raise LicenciaError("Correo o contraseña incorrectos.")
    raise LicenciaError(f"No se pudo iniciar sesión ({st}). {msg}".strip())


def login_google(timeout=180):
    """Login con Google: abre el navegador y espera el callback local (PKCE)."""
    import threading, webbrowser
    from http.server import HTTPServer, BaseHTTPRequestHandler
    from urllib.parse import urlencode, urlparse, parse_qs

    verifier = base64.urlsafe_b64encode(os.urandom(40)).decode().rstrip("=")
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")

    resultado, listo = {}, __import__("threading").Event()

    class _H(BaseHTTPRequestHandler):
        def do_GET(self):
            q = parse_qs(urlparse(self.path).query)
            code = (q.get("code") or [""])[0]
            ok = bool(code)
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            cuerpo = ("<h2>Listo ✔</h2><p>Ya puedes volver a Gestiq.</p>" if ok else
                      "<h2>Algo salió mal</h2><p>Vuelve a Gestiq e intenta de nuevo.</p>")
            self.wfile.write((
                "<html><body style='font-family:sans-serif;background:#15151F;"
                "color:#EDEDF5;text-align:center;padding-top:80px'>"
                + cuerpo + "</body></html>").encode())
            resultado["code"] = code
            listo.set()
        def log_message(self, *a):  # silencio
            pass

    srv = HTTPServer(("127.0.0.1", 0), _H)
    puerto = srv.server_address[1]
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        params = urlencode({"provider": "google",
                            "redirect_to": f"http://localhost:{puerto}",
                            "code_challenge": challenge,
                            "code_challenge_method": "s256"})
        webbrowser.open(f"{SUPABASE_URL}/auth/v1/authorize?{params}")
        if not listo.wait(timeout):
            raise LicenciaError("Tiempo de espera agotado. Intenta de nuevo.")
    finally:
        threading.Thread(target=srv.shutdown, daemon=True).start()
    code = resultado.get("code", "")
    if not code:
        raise LicenciaError("Google no autorizó el acceso.")
    st, j = _post("/auth/v1/token?grant_type=pkce",
                  {"auth_code": code, "code_verifier": verifier})
    if st == 200:
        return _sesion_desde(j)
    raise LicenciaError(f"No se pudo completar el acceso con Google ({st}).")


def restaurar_sesion():
    """Reanuda la sesión guardada sin pedir contraseña. None si no hay."""
    rt = _leer("refresh")
    if not rt:
        return None
    st, j = _post("/auth/v1/token?grant_type=refresh_token", {"refresh_token": rt})
    if st == 200:
        return _sesion_desde(j)
    _borrar("refresh")          # token revocado/expirado → pedirá login
    return None


def cerrar_sesion():
    _borrar("refresh")


def ultimo_email():
    return _leer("email") or ""


# ── Verificación de licencia (servidor decide todo) ──────────────────────────
MOTIVOS = {
    "SIN_LICENCIA": "Tu cuenta no tiene una licencia asignada.\nEscríbenos para activarla.",
    "OTRO_EQUIPO":  "Tu licencia ya está activada en otro equipo.\nEscríbenos si cambiaste de computador.",
    "SUSPENDIDA":   "Tu licencia está suspendida.\nEscríbenos para reactivarla.",
    "VENCIDA":      "Tu suscripción venció.\nRenueva para seguir usando Gestiq.",
}

def motivo(r):
    return MOTIVOS.get((r or {}).get("reason", ""), "Licencia no válida.")


def verificar(sesion):
    """Valida la licencia en el servidor. Renueva el token si expiró.
    Devuelve el dict del servidor: {ok, expires_at, server_time, ...}."""
    st, j = _post("/rest/v1/rpc/check_license",
                  {"p_machine_id": machine_id()}, token=sesion["access_token"])
    if st == 401:               # access token vencido → renovar y reintentar una vez
        nueva = restaurar_sesion()
        if not nueva:
            raise LicenciaError("La sesión expiró. Inicia sesión de nuevo.")
        sesion.update(nueva)
        st, j = _post("/rest/v1/rpc/check_license",
                      {"p_machine_id": machine_id()}, token=sesion["access_token"])
    if st != 200 or not isinstance(j, dict):
        raise LicenciaError(f"El servidor de licencias respondió {st}.")
    return j


# ── Perfil de la cuenta (nombre, foto, tema, módulo) en el servidor ──────────
_PERFIL_COLS = ("nombre", "foto", "tema", "modulo")

def _uid(sesion):
    """user id (sub) del JWT, sin verificar firma."""
    try:
        p = (sesion or {}).get("access_token", "").split(".")[1]
        p += "=" * (-len(p) % 4)
        return json.loads(base64.urlsafe_b64decode(p)).get("sub", "")
    except Exception:
        return ""

def get_profile(sesion):
    """Lee el perfil de la cuenta desde el servidor. {} si no hay o falla."""
    try:
        req = urllib.request.Request(
            SUPABASE_URL.rstrip("/")
            + "/rest/v1/profiles?select=nombre,foto,tema,modulo",
            headers={"apikey": SUPABASE_ANON_KEY,
                     "Authorization": f"Bearer {sesion['access_token']}"},
            method="GET")
        with urllib.request.urlopen(req, timeout=TIMEOUT, context=_ssl_ctx()) as r:
            arr = json.loads(r.read().decode() or "[]")
        if isinstance(arr, list) and arr:
            return {k: v for k, v in arr[0].items() if v is not None}
    except Exception:
        pass
    return {}

def set_profile(sesion, **kw):
    """Crea/actualiza (upsert) el perfil de la cuenta en el servidor."""
    uid = _uid(sesion)
    if not uid:
        return False
    row = {"user_id": uid}
    for k in _PERFIL_COLS:
        if k in kw:
            row[k] = kw[k]
    try:
        req = urllib.request.Request(
            SUPABASE_URL.rstrip("/") + "/rest/v1/profiles?on_conflict=user_id",
            data=json.dumps(row).encode(),
            headers={"apikey": SUPABASE_ANON_KEY,
                     "Authorization": f"Bearer {sesion['access_token']}",
                     "Content-Type": "application/json",
                     "Prefer": "resolution=merge-duplicates,return=minimal"},
            method="POST")
        with urllib.request.urlopen(req, timeout=TIMEOUT, context=_ssl_ctx()) as r:
            return r.status in (200, 201, 204)
    except Exception:
        return False
