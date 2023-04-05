import logging
import os


# --- Début des variables ---
dir = os.path.dirname(__file__)
dirLog = dir + "\\..\\"
LogFile = os.path.join(dirLog, 'ExportAD.log')
# --- Fin des variables ---


# --- Début des fonctions --
def setup_logging():
    """Setup et configure un fichier de log."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(name)s - %(message)s',
        filename=LogFile,
        filemode='a',
        encoding='utf-8'
        )
# --- Fin des fonctions ---