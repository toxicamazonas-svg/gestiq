# -*- coding: utf-8 -*-
"""Fuente única de la versión de Gestiq.

Al publicar una versión nueva, cambiar SOLO aquí (semver v1.0.X, no saltar).
La usan gestiq.py, gestiq_web.py, updater.py y la UI (vía api.version()).
El auto-update compara este valor con el último tag publicado en GitHub.
"""
VERSION = "1.0.11"
