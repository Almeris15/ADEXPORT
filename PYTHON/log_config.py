import logging
import os


# --- Start of variables ---
dir = os.path.dirname(__file__)
dirLog = dir + "\\..\\"
LogFile = os.path.join(dirLog, 'ExportAD.log')
# --- End of variables ---


# --- Start of functions --
def setup_logging():
    """Setup and configure a log file."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(name)s - %(message)s',
        filename=LogFile,
        filemode='a',
        encoding='utf-8'
        )
# --- End of functions ---