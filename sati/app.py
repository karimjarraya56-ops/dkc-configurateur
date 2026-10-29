"""Configurateur SATI Tunisia – même outil que DKC, habillage SATI.
Lancer :  streamlit run sati/app.py
"""
import runpy
import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

runpy.run_path(str(RACINE / "app.py"), init_globals={"MARQUE": "sati"}, run_name="__main__")
