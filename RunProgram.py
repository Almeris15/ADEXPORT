#NAME : RunProgram.py
#Author : Almeris15
#Description :
# Script pour executer l'ensemble des scripts
#  
# Changelog:
# 1.0.0 - Initial release

# Imports nécessaires
import argparse
import subprocess
import os
import datetime

# --- Début des variables ---
dir = os.path.dirname(__file__) # import absolute path
# Path pour les script et programmes
program1 = os.path.join(dir, ".\\PYTHON\\CSVtoMatriceExcel.py") 
program2 = os.path.join(dir, ".\\PYTHON\\editExcelFile.py")
program3 = os.path.join(dir, ".\\PYTHON\\Compare.py")
# --- Fin des variables ---

# --- Début des fonctions ---
# Fonction pour lancer les script Python
def RunProgram(program):
    print("run du programme ", program)
    subprocess.run(["python", program, output_file_name]) # Lancement du script avec options

# Fonction pour supprimer les csv
def Suprimme(csv):
    try:
        os.remove(csv)
        print("supression de ", csv)
    except FileNotFoundError:
        pass
# --- Fin des fonctions ---
 
parser = argparse.ArgumentParser() # Initialize the argument parser
parser.add_argument("--Create", help="Name of the output file") # Add an option for the output file
parser.add_argument("--Compare", help="Name of the output file") # Add an option for the output file
args = parser.parse_args() # Parse the command-line arguments

date = datetime.datetime.now().strftime("%Y%m%d") # Get date avec la forme YYYYMMDD

if args.Create:
    output_file_name = args.Create
    print(f"Creation d'une matrice. Nom du fichier de sortie :'{output_file_name}_{date}.xlsx'")
    RunProgram(program1) # Lancement du programme CSVToMatrice
    RunProgram(program2) # Lancement du programme EditExcel

elif args.Compare:
    output_file_name = args.Compare
    print(f"Comparaison de 2 matrices.")
    RunProgram(program3) # Lancement du programme python pour Comparer 2 matrices

else:
    print("Veuillez choisir une option")

print("Program Finish")
