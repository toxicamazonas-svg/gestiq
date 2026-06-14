#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gestiq v1.0.9 — interfaz liquid glass (pywebview).
La lógica del bot, Excel y licencias se reutiliza COMPLETA de gestiq.py;
este archivo solo cambia la capa visual (HTML/CSS real con blur).
Requiere: pip install pywebview  (además de lo de siempre)
"""

import os, sys, json, threading, time, webbrowser
from datetime import datetime

import webview                    # pywebview
import openpyxl

import gestiq as G                # ← toda la lógica existente
licencia = G.licencia
from version import VERSION
import updater


# ── Sustitutos sin Tk ────────────────────────────────────────────────────────
class FakeVar:
    """Imita ctk.StringVar (.get/.set)."""
    def __init__(self, v=""): self.v = v
    def get(self): return self.v
    def set(self, v): self.v = v


class Dummy:
    """Imita cualquier widget: acepta .configure(...) y poco más."""
    def configure(self, **k): pass
    def winfo_exists(self): return False
    def set(self, *a): pass


class MsgShim:
    """Sustituye tkinter.messagebox dentro de gestiq → toasts en la web."""
    def __init__(self, api): self.api = api
    def showinfo(self, t, m, **k):    self.api.js_toast("ok", t, m)
    def showwarning(self, t, m, **k): self.api.js_toast("warn", t, m)
    def showerror(self, t, m, **k):   self.api.js_toast("err", t, m)
    def askyesno(self, t, m, **k):    return True


# ── Pestañas "headless": lógica de gestiq.py sin widgets ────────────────────
class TabWeb:
    """Mixin que reemplaza la parte Tk de BaseTab. El MRO hace que estos
    métodos tapen a los de BaseTab; _automate y los helpers se heredan."""

    def init_comun(self, api, key):
        self.api = api
        self.key = key
        self.app = api                      # api.log(tab, msg, lvl) como App.log
        self.xl_path = None
        self.wb = None
        self._stop = False
        self._running = False
        self._loop = self._task = self._thread = None
        self._login_ev = None
        self._login_dlg = None
        self.v_sheet = FakeVar()
        # widgets que tocan los métodos heredados (_do_stop, etc.)
        self.btn_start = self.btn_stop = self.btn_download = Dummy()
        self.lbl_prog = self.lbl_pct = self.prog = self.cb_sheet = Dummy()

    # gestiq llama self.after(0, fn) desde el hilo del bot: aquí no hay loop
    # de Tk, así que se ejecuta directo (evaluate_js es thread-safe).
    def after(self, _ms, fn=None, *a):
        if fn is None: return
        try: fn(*a)
        except Exception as e:
            self.api.log(self, f"Error de interfaz: {e}", "error")

    def after_cancel(self, *_): pass

    def _set_running(self, r): self._running = r

    def _upd(self, cur, tot, msg=""):
        self.api.js(f"G.progreso({json.dumps(self.key)},{int(cur)},{int(tot)},{json.dumps(str(msg))})")

    def _set_prog(self, v):
        self.api.js(f"G.barra({json.dumps(self.key)},{float(v)})")

    def _on_finish(self):
        self._set_running(False)
        self.api.js(f"G.fin({json.dumps(self.key)},{json.dumps(bool(self._stop))})")

    def _ask_ready(self, event, system_name):
        self._login_ev = event
        self.api.js(f"G.pedirLogin({json.dumps(self.key)},{json.dumps(system_name)})")


class ImagineWeb(TabWeb, G.ImagineTab):
    def __init__(self, api):
        self.init_comun(api, "imagine")
        self.v_caso = FakeVar("CASO")
        self.v_out  = FakeVar("IMAGINE")


class GuardianWeb(TabWeb, G.GuardianTab):
    def __init__(self, api):
        self.init_comun(api, "guardian")
        self.v_cron = FakeVar("CRONOGRAMA")
        self.v_sec  = FakeVar("SECUENCIA")
        self.v_out  = FakeVar("GUARDIAN")


# ── API expuesta a JavaScript ────────────────────────────────────────────────
class Api:
    def __init__(self):
        self.win = None
        self._lic = None
        self._pend2fa = None
        self._cid2fa = ""
        self._enrol = None
        self._enrolCid = ""
        self._plan = "completo"
        self._hb_on = False
        self.tabs = {"imagine": ImagineWeb(self), "guardian": GuardianWeb(self)}
        G.messagebox = MsgShim(self)        # los messagebox del bot → toasts

    # ── Puente Python → JS ──
    def js(self, code):
        try:
            if self.win: self.win.evaluate_js(code)
        except Exception:
            pass

    def js_toast(self, tipo, titulo, msg):
        self.js(f"G.toast({json.dumps(tipo)},{json.dumps(str(titulo))},{json.dumps(str(msg))})")

    def log(self, tab, msg, lvl="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.js(f"G.log({json.dumps(tab.key)},{json.dumps(ts)},{json.dumps(str(msg))},{json.dumps(lvl)})")

    # ── Licencia ──
    def _lic_dict(self, modo, **k):
        d = {"modo": modo}
        d.update(k)
        return d

    def estado_inicial(self):
        if licencia is None or not licencia.configurado():
            return self._lic_dict("bloqueado", msg="Error interno de licencias. "
                                  "Reinstala la aplicación o contacta soporte.")
        try:
            s = licencia.restaurar_sesion()
            if s is None:
                em = licencia.ultimo_email()
                return self._lic_dict("login", email=em,
                                      prefs=G._prefs_get(em))
            r = licencia.verificar(s)
            if r.get("ok"):
                return self._lic_ok(s, r)
            return self._lic_dict("bloqueado", msg=licencia.motivo(r))
        except Exception as e:
            return self._lic_dict("bloqueado", msg=str(e), reintentar=True)

    def _tras_sesion(self, s, em=""):
        r = licencia.verificar(s)
        if r.get("ok"):
            return self._lic_ok(s, r)
        licencia.cerrar_sesion()
        return self._lic_dict("bloqueado", msg=licencia.motivo(r))

    def login(self, em, pw):
        try:
            s = licencia.login(em, pw)
            if s.get("requiere_2fa"):
                self._pend2fa = s
                return self._lic_dict("2fa", email=s.get("email", ""))
            return self._tras_sesion(s, em)
        except Exception as e:
            return self._lic_dict("login", msg=str(e), email=em)

    def reto_2fa(self):
        """Prepara el reto del 2do factor; en SMS, envía el código."""
        try:
            s = self._pend2fa or {}
            self._cid2fa = licencia.challenge_2fa(s.get("access_token", ""), s.get("factor_id", ""))
            return {"ok": True, "tipo": s.get("factor_tipo", "totp")}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def verificar_2fa(self, code):
        try:
            s = self._pend2fa or {}
            ses = licencia.verify_2fa(s.get("access_token", ""), s.get("factor_id", ""),
                                      self._cid2fa, code)
            self._pend2fa = None; self._cid2fa = ""
            return self._tras_sesion(ses)
        except Exception as e:
            return self._lic_dict("2fa", msg=str(e),
                                  email=(self._pend2fa or {}).get("email", ""),
                                  tipo=(self._pend2fa or {}).get("factor_tipo", "totp"))

    def login_google(self):
        try:
            return self._tras_sesion(licencia.login_google())
        except Exception as e:
            return self._lic_dict("login", msg=str(e))

    def activar_2fa(self, tipo="totp", phone=""):
        try:
            d = licencia.enrolar_2fa(self._lic, tipo, phone)
            self._enrol = {"factor_id": d["factor_id"], "tipo": tipo}
            self._enrolCid = ""
            return {"ok": True, **d}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def enviar_codigo_2fa(self):
        """Para SMS en alta: envía el código al teléfono del factor recién creado."""
        try:
            self._enrolCid = licencia.challenge_2fa(self._lic.get("access_token", ""),
                                                    (self._enrol or {}).get("factor_id", ""))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def confirmar_2fa(self, code):
        try:
            fid = (self._enrol or {}).get("factor_id", "")
            cid = self._enrolCid or licencia.challenge_2fa(self._lic.get("access_token", ""), fid)
            self._lic = licencia.verify_2fa(self._lic.get("access_token", ""), fid, cid, code)
            self._enrol = None; self._enrolCid = ""
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def desactivar_2fa(self, factor_id):
        try:
            licencia.desactivar_2fa(self._lic, factor_id)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def _lic_ok(self, s, r):
        self._lic = s
        self._plan = str((r or {}).get("plan") or "completo").lower()
        if self._plan not in ("imagine", "guardian", "completo"):
            self._plan = "completo"
        self._arrancar_heartbeat()
        return self._lic_dict("ok", email=s.get("email", ""), plan=self._plan,
                              prefs=self._perfil())

    def salir(self):
        for t in self.tabs.values():
            if t._running: t._do_stop()
        self._lic = None
        if licencia:
            try: licencia.cerrar_sesion()
            except Exception: pass
        em = licencia.ultimo_email() if licencia else ""
        return self._lic_dict("login", email=em, prefs=G._prefs_get(em))

    def _arrancar_heartbeat(self):
        if self._hb_on: return
        self._hb_on = True
        def bucle():
            while True:
                time.sleep(600)
                if self._lic is None: continue
                try:
                    r = licencia.verificar(self._lic)
                    if r.get("ok"):
                        self._plan = str(r.get("plan") or self._plan).lower()
                        self.js(f"G.plan({json.dumps(self._plan)})")
                    else:
                        self._bloquear(licencia.motivo(r))
                except Exception as e:
                    self._bloquear(str(e), True)
        threading.Thread(target=bucle, daemon=True).start()

    def _bloquear(self, msg, reintentar=False):
        for t in self.tabs.values():
            if t._running: t._do_stop()
        self._lic = None
        self.js(f"G.lic({json.dumps(self._lic_dict('bloqueado', msg=msg, reintentar=reintentar))})")

    # ── Archivo ──
    def elegir_archivo(self, m):
        tab = self.tabs[m]
        sel = self.win.create_file_dialog(
            webview.OPEN_DIALOG, file_types=("Excel (*.xlsx;*.xls)", "Todos (*.*)"))
        if not sel: return None
        p = sel[0] if isinstance(sel, (list, tuple)) else sel
        try:
            wb = openpyxl.load_workbook(p)
        except Exception as e:
            self.log(tab, f"Error al abrir archivo: {e}", "error")
            return {"error": "El archivo no se pudo abrir. Ciérralo en Excel e inténtalo de nuevo."}
        tab.xl_path, tab.wb = p, wb
        nombre = os.path.basename(p)
        self.log(tab, f"Cargado: {nombre}  ({len(wb.sheetnames)} hoja(s))", "ok")
        return {"nombre": nombre, "hojas": wb.sheetnames}

    def guardar_copia(self, m):
        tab = self.tabs[m]
        if not tab.wb: return {"error": "No hay resultados que guardar."}
        sel = self.win.create_file_dialog(
            webview.SAVE_DIALOG, file_types=("Excel (*.xlsx)",),
            save_filename=f"resultado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not sel: return None
        p = sel[0] if isinstance(sel, (list, tuple)) else sel
        try:
            tab.wb.save(p)
            self.log(tab, f"Excel guardado: {os.path.basename(p)}", "ok")
            return {"ruta": p}
        except Exception as e:
            self.log(tab, f"Error al guardar: {e}", "error")
            return {"error": "Revisa que el archivo de destino no esté abierto en Excel."}

    # ── Ejecución ──
    def iniciar(self, m, cfg):
        tab = self.tabs[m]
        if tab._running: return {"error": "Ya hay una consulta en curso."}
        if not tab.xl_path: return {"error": "Primero selecciona un archivo Excel."}
        if not (cfg or {}).get("hoja"): return {"error": "Selecciona la hoja del Excel."}
        if not G.HAVE_PW: return {"error": "Playwright no está instalado en este equipo."}

        # Licencia obligatoria antes de cada ejecución (igual que lic_check_run)
        if licencia is None or not licencia.configurado():
            return {"error": "Error interno de licencias."}
        if self._lic is None:
            self.js("G.lic({\"modo\":\"login\"})")
            return {"error": "Inicia sesión para continuar."}
        try:
            r = licencia.verificar(self._lic)
            if not r.get("ok"):
                self._bloquear(licencia.motivo(r))
                return {"error": "Licencia no válida."}
            self._plan = str(r.get("plan") or self._plan).lower()
            self.js(f"G.plan({json.dumps(self._plan)})")
            if self._plan != "completo" and self._plan != m:
                return {"error": "Tu plan actual no incluye este módulo."}
        except Exception as e:
            return {"error": f"No se pudo validar la licencia: {e}"}

        try:
            tab.wb = openpyxl.load_workbook(tab.xl_path)
        except Exception as e:
            self.log(tab, f"Error al recargar archivo: {e}", "error")
            return {"error": "El archivo no se pudo recargar. Ciérralo en Excel e inténtalo de nuevo."}

        tab.v_sheet.set(cfg["hoja"])
        for k, v in (cfg or {}).items():
            if k == "hoja" or not str(v).strip(): continue
            var = getattr(tab, "v_" + k, None)
            if var: var.set(str(v).strip())

        tab._stop = False
        tab._login_ev = None
        tab._set_running(True)
        tab._thread = threading.Thread(target=tab._runner, daemon=True)
        tab._thread.start()

        def vigilar(t=tab):
            t._thread.join()
            t._thread = None
            t._on_finish()
        threading.Thread(target=vigilar, daemon=True).start()
        return {"ok": True}

    def detener(self, m):
        tab = self.tabs[m]
        if tab._running: tab._do_stop()
        return {"ok": True}

    def continuar_login(self, m):
        ev = self.tabs[m]._login_ev
        if ev and not ev.is_set(): ev.set()
        return {"ok": True}

    def cancelar_login(self, m):
        tab = self.tabs[m]
        tab._stop = True
        ev = tab._login_ev
        if ev and not ev.is_set(): ev.set()
        return {"ok": True}

    # ── Perfil y preferencias (asociadas a la cuenta) ──
    def _email(self):
        return (self._lic or {}).get("email", "")

    def _perfil(self):
        """Perfil de la cuenta: el servidor manda, con respaldo local.
        La primera vez tras actualizar, sube el perfil local al servidor."""
        em = self._email()
        local = G._prefs_get(em)
        if self._lic is None or licencia is None:
            return local
        try:
            serv = licencia.get_profile(self._lic)
        except Exception:
            serv = {}
        if serv:
            if serv != local:
                try: G._prefs_set(em, **serv)
                except Exception: pass
            return serv
        if local:                      # servidor vacío → migrar lo local
            try: licencia.set_profile(self._lic, **local)
            except Exception: pass
        return local

    def prefs_get(self):
        return self._perfil()

    def prefs_set(self, kw):
        em = self._email()
        if not em:
            return {"error": "No hay sesión activa."}
        try:
            kw = {k: v for k, v in dict(kw or {}).items()
                  if k in ("nombre", "foto", "tema", "modulo")}
            G._prefs_set(em, **kw)
            if self._lic is not None and licencia is not None:
                try: licencia.set_profile(self._lic, **kw)
                except Exception: pass
            return {"ok": True, "prefs": self._perfil()}
        except Exception as e:
            return {"error": str(e)}

    def leer_foto(self):
        """Diálogo de imagen; devuelve base64 sin procesar (recorte en JS)."""
        sel = self.win.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=("Imágenes (*.png;*.jpg;*.jpeg;*.webp;*.gif;*.bmp)",))
        if not sel:
            return None
        p = sel[0] if isinstance(sel, (list, tuple)) else sel
        try:
            if os.path.getsize(p) > 12_000_000:
                return {"error": "La imagen pesa demasiado (máx. 12 MB)."}
            import base64
            with open(p, "rb") as f:
                raw = base64.b64encode(f.read()).decode()
            ext = os.path.splitext(p)[1].lower().lstrip(".") or "png"
            if ext == "jpg":
                ext = "jpeg"
            return {"b64": raw, "mime": f"image/{ext}"}
        except Exception as e:
            return {"error": str(e)}

    # ── Varios ──
    def abrir_web(self):
        webbrowser.open(G.REGISTRO_URL)
        return {"ok": True}

    def version(self):
        return VERSION


# ── Arranque ─────────────────────────────────────────────────────────────────
def _ruta(nombre):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nombre)


def main():
    api = Api()
    kwargs = dict(
        title="Gestiq", url=_ruta("gestiq_ui.html"), js_api=api,
        width=1140, height=780, min_size=(960, 660),
        background_color="#15151F",
    )
    try:
        api.win = webview.create_window(vibrancy=True, **kwargs)   # blur nativo (solo macOS)
    except TypeError:
        api.win = webview.create_window(**kwargs)

    def al_cerrar():
        for t in api.tabs.values():
            if t._running:
                try: t._do_stop()
                except Exception: pass
        try: updater.aplicar_y_reiniciar()   # si hay update listo, reemplaza y relanza
        except Exception: pass
        return True
    api.win.events.closing += al_cerrar

    # Auto-actualización silenciosa en segundo plano.
    def _aviso_update(tag):
        api.js_toast("ok", "Actualización",
                     f"Gestiq {tag} se instalará al cerrar la aplicación.")
    try:
        updater.iniciar_en_segundo_plano(on_listo=_aviso_update)
    except Exception:
        pass

    try:
        webview.start(private_mode=False)   # conserva el tema elegido
    except TypeError:
        webview.start()


if __name__ == "__main__":
    main()
