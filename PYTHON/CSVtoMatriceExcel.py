import logging
import sys
import os
import pandas as pd
from datetime import datetime
from log_config import setup_logging


# --- variables ---
dir = os.path.dirname(__file__)  # Emplacement absolue du fichier
dirExcel = dir + "\\..\\ExportExcel\\"  # Emplacement du dossier Excel
OutputFile = sys.argv[1]  # Nom du fichier Excel en argv d'appel de script
dirCSV = dir + "\\..\\ImportCSV\\"  # Emplacement du dossier CSV
# --- Fin des variables ---


# --- Début des fonctions ---
def FileDate(fileName):
    """Fonction pour choisir le fichier par date."""
    latest_file_name = None
    latest_file_date = 0
    file_date_str = fileName.split("_")[-1].split(".")[0]  # Présume que le format est "ADmembers_'Ad.Name'_YYYYMMDD.csv" et Garde que la date
    file_date = int(file_date_str)
    if latest_file_date is None or file_date > latest_file_date:  # replace file by the latest
        latest_file_name = os.path.join(dirCSV, fileName)
        latest_file_date = file_date
    return latest_file_name  # renvoi le fichier le plus récent


def Call_files():
    """Fonction pour appel de fichier."""
    for file_name in os.listdir(dirCSV):  # se place dans le dossier Temp
        if file_name.startswith("ADmembers_") and file_name.endswith(".csv"):  # ne choisit que les fichiers csv commençant par Admembers_
            User_file = FileDate(file_name)
        if file_name.startswith("ADgroups_") and file_name.endswith(".csv"):  # ne choisit que les fichiers csv commençant par Adroups_
            Group_file = FileDate(file_name)
        if file_name.startswith("ADmembership_") and file_name.endswith(".csv"):  # ne choisit que les fichiers csv commençant par Admembership_
            Membership_file = FileDate(file_name)
    return User_file, Group_file, Membership_file  # return les 3 fichier csv nécessaires


def File_to_list(filecsv):
    """Fonction pour les 2 premier fichiers csv : User et Groupes."""
    with open(filecsv, 'r') as fileR:
        List = fileR.readlines()[1:]  # read all lines except the first one
    ListFinal = [i.replace('"', '').replace("\n", '') for i in List]  # Structure le code en enlevant les quotes et \n
    return ListFinal


def Group_Membership(filecsv):
    """Fonction pour le 3ème fichier csv avec groupemembership."""
    with open(filecsv, 'r') as fileR:
        lineR = fileR.readlines()[1:]  # read all lines except the first one
    dico = {}
    for line in lineR:
        name, group = map(str.strip, line.strip().split(","))  # Sépare la liste csv en 2 colonnes à ","
        name = name.strip('"')
        group = group.strip('"')
        dico.setdefault(group, []).append(name)  # append the new number to the existing array else create a new array
    return dico


def Generate_table(list1, list2, dico):
    """Fonction pour générer le tableau."""
    tableau = []
    for name in list2:  # Pour chaque User
        row = ["X" if name in dico.get(i, []) else " " for i in list1]  # Pour chaque ligne met des X si membership sinon ne met rien
        tableau.append(row)  # crée la ligne
    table = pd.DataFrame(tableau, columns=list1, index=list2)  # crée le tableau
    return table
# --- Fin des fonctions ---


# --- Appel aux fonctions ---
setup_logging()
User, Groups, GroupMembership = Call_files()

ListUserFinal = File_to_list(User)
ListGroupFinal = File_to_list(Groups)
dictionnaire = Group_Membership(GroupMembership)
table = Generate_table(ListGroupFinal, ListUserFinal, dictionnaire)

# Définition du nom du fichier excel en sortie et de son emplacement
date = datetime.now().strftime("%Y%m%d")  # Prend la date du jour
ExcelName = OutputFile + "_" + date + ".xlsx"  # Construit le nom de fichier de sortie
ExcelFile = os.path.join(dirExcel, ExcelName)  # Join la path et le nom de fichier

writer = pd.ExcelWriter(ExcelFile)  # creating excel writer object
table.to_excel(writer)  # write dataframe to excel
writer.close()  # close the excel

print("Excel File is successfully created")  # Print excel created
logging.info('Created File Excel : %s', ExcelName) # log
