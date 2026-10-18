"""Lógica de los cuadres de créditos MG Colima."""
try:                      # hora de México en todo lo que use datetime.now() / date.today()
    import zona  # noqa: F401
except ImportError:       # uso suelto de cuadres/ (Colab) sin la raíz del repo
    import os as _os, time as _time
    _os.environ["TZ"] = "America/Mexico_City"
    if hasattr(_time, "tzset"):
        _time.tzset()
