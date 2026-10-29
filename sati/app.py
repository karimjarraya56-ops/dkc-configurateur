"""Configurateur SATI Tunisia – même outil que DKC, habillage SATI.
Lancer :  streamlit run sati/app.py
"""
import runpy
import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

# Streamlit ne recharge que les modules du dossier sati/ : on oublie ceux de la racine
# pour que les modifications de i18n.py, catalogue_s5.py, export_liste.py soient prises en compte
for module in ("i18n", "catalogue_s5", "export_liste"):
    sys.modules.pop(module, None)

runpy.run_path(str(RACINE / "app.py"), init_globals={"MARQUE": "sati"}, run_name="__main__")
