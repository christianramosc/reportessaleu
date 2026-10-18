# -*- coding: utf-8 -*-
"""Hora de México para toda la app.

Streamlit Cloud corre en UTC: a partir de las 6 pm de Colima (UTC-6) datetime.now() y date.today() ya daban el día
siguiente. Al importar este módulo el proceso usa la zona America/Mexico_City, así todas las fechas y horas
(«hoy», el Historial, los nombres de archivo, «Generado: …») salen en hora local.
"""
import os
import time

ZONA = "America/Mexico_City"


def fijar():
    if os.environ.get("TZ") != ZONA:
        os.environ["TZ"] = ZONA
        if hasattr(time, "tzset"):
            time.tzset()


fijar()
