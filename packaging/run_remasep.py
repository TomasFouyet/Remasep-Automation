"""Punto de entrada para PyInstaller (Sprint 3.11).

Bootstrap fino a propósito: vive **fuera** del paquete ``remasep`` para que
PyInstaller no lo trate como parte del paquete (evitaría que ``remasep.main``
se importe dos veces bajo dos nombres distintos). Toda la lógica real vive en
``remasep.main`` / ``remasep.ui.main_window`` — igual que ejecutando
``python -m remasep.main`` en desarrollo.
"""

from remasep.main import main

if __name__ == "__main__":
    main()
