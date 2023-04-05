import argparse
import subprocess
import os
import datetime
import sys
import logging
from PYTHON.log_config import setup_logging

# --- Début des variables ---
dir = os.path.dirname(__file__)  # import absolute path
# Path pour les script et programmes
program1 = os.path.join(dir, ".\\PYTHON\\CSVtoMatriceExcel.py")
program2 = os.path.join(dir, ".\\PYTHON\\editExcelFile.py")
program3 = os.path.join(dir, ".\\PYTHON\\Compare.py")
# --- Fin des variables ---


# --- Début des fonctions ---
def RunProgram(program):
    """Fonction pour lancer les script Python."""
    try:
        print("run du programme ", program)
        logging.info('lancement du script : %s', program)
        subprocess.run(["python", program, output_file_name])  # Lancement du script avec options
    except FileNotFoundError:
        print(f"Script {program} not found")
        logging.error('Script %s non trouve', program)


def DeleteFile(csv):
    """Fonction pour supprimer les csv."""
    try:
        os.remove(csv)
        print("supression de ", csv)
    except FileNotFoundError:
        pass


def install(package):
    """Fonctions pour installer les packages nécessaires manquant."""
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
# --- Fin des fonctions ---


# --- Appel aux fonctions ---
install('pandas')
install('openpyxl')

setup_logging()

parser = argparse.ArgumentParser(description='Process csv into matrix excel')  # Initialize the argument parser
parser.add_argument("--Create", help="Cree une matrice grâce aux csv", type=str, nargs=1)  # Add an option for the program
parser.add_argument("--Compare", help="Compare 2 matrice grâce aux csv", type=str, nargs=1)  # Add an option for the program

try:
    args = parser.parse_args()  # Parse the command-line arguments
except Exception as e:
    logging.error("Error command")

date = datetime.datetime.now().strftime("%Y%m%d")  # Get date avec la forme YYYYMMDD

if args.Create:
    output_file_name = args.Create[0]
    print(f"Creation d'une matrice. Nom du fichier de sortie :'{output_file_name}_{date}.xlsx'")
    logging.info("Creation of a matrix. Name of the file in exit :'%s_%s.xlsx'", output_file_name, date)
    RunProgram(program1)  # Lancement du programme CSVToMatrice
    RunProgram(program2)  # Lancement du programme EditExcel
elif args.Compare:
    output_file_name = args.Compare[0]
    print(f"Comparaison de 2 matrices.")
    logging.info("Compare 2 matrix")
    RunProgram(program3)  # Lancement du programme python pour Comparer 2 matrices
else:
    print("Veuillez choisir une option")
    logging.error("Error no option has been chosen")

print("Program Finish")
logging.info('End of program')
