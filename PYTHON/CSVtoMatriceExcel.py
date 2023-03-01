#NAME : CSVtoMatriceExcel.py
#Author : Almeris15
#Description : 
#  
# Transformer la sortie des informations AD de liste à dict pour traitement.
# les csv sont transformes en une matrice dans un fichier excel.
#
# Pour ce programme nous avons besoin de openpyxl afin de mettre en page le fichier excel de sortie
# pandas est également requis
# L'installation de ce module est prévu dans le script
# 
# Pensez à mettre le nom du fichier voulu en argument
#
# Changelog:
# 1.0.0 - Initial release

import sys
import os 
import datetime
import pandas as pd

# --- variables ---
dir = os.path.dirname(__file__) # Emplacement du fichier dans l'ordinateur/serveur
dirExcel = dir + "\\..\\EXCEL_FILES\\"
OutputFile = sys.argv[1] # Nom du fichier Excel = 1er argument en appel de script
dirCSV = dir + "\\..\\TEMP\\" # Emplacement du dossier Temp
# --- Fin des variables ---

# --- Début des fonctions ---
# Fonction pour choisir le fichier par date
def FileDate(fileName):
    latest_file_name = None
    latest_file_date = 0
    file_date_str = fileName.split("_")[-1].split(".")[0]  # Présume que le format est "ADmembers_'Ad.Name'_YYYYMMDD.csv" et Garde que la date
    file_date = int(file_date_str)
    if latest_file_date is None or file_date > latest_file_date: # replace file by the latest
        latest_file_name = os.path.join(dirCSV, fileName)
        latest_file_date = file_date
    return latest_file_name # renvoi le fichier le plus récent

# Fonction pour appel de fichier
def Call_files():
    for file_name in os.listdir(".\\TEMP\\"): # se place dans le dossier Temp
        if file_name.startswith("ADmembers_") and file_name.endswith(".csv"): # ne choisit que les fichiers csv commençant par Admembers_
            User_file = FileDate(file_name)
        if file_name.startswith("ADgroups_") and file_name.endswith(".csv"): # ne choisit que les fichiers csv commençant par Adroups_
            Group_file = FileDate(file_name)
        if file_name.startswith("ADmembership_") and file_name.endswith(".csv"): # ne choisit que les fichiers csv commençant par Admembership_
            Membership_file = FileDate(file_name)
    return User_file, Group_file, Membership_file # return les 3 fichier csv nécessaires

# Fonction pour les 2 premier fichiers csv avec les User et les Groupes
def File_to_list(filecsv):
    with open(filecsv, 'r') as fileR:
        List = fileR.readlines()[1:] # read all lines except the first one
    ListFinal = [i.replace('"','').replace("\n",'') for i in List] # Structure le code en enlevant les quotes et \n
    return ListFinal

# Fonction pour le 3ème fichier csv avec groupemembership 
def Group_Membership(filecsv):
    with open(filecsv, 'r') as fileR:
        lineR = fileR.readlines()[1:] # read all lines except the first one
    dico = {}
    for line in lineR:
        name, group = map(str.strip, line.strip().split(",")) # Sépare la liste csv en 2 colonnes à ","
        name = name.strip('"')
        group = group.strip('"')
        dico.setdefault(group, []).append(name) # append the new number to the existing array else create a new array
    return dico

#fonction pour générer le tableau
def Generate_table(list1, list2, dico):
    tableau = []
    for name in list2: # Pour chaque USer 
        row = ["X" if name in dico.get(i, []) else " " for i in list1] # Pour chaque ligne met des X si membership sinon ne met rien
        tableau.append(row) # crée la ligne
    table = pd.DataFrame(tableau, columns=list1, index=list2) # crée le tableau
    return table
# --- Fin des fonctions ---

# Appel aux fonctions
User, Groups, GroupMembership = Call_files()

ListUserFinal = File_to_list(User)
ListGroupFinal = File_to_list(Groups)
dictionnaire = Group_Membership(GroupMembership)
table = Generate_table(ListGroupFinal, ListUserFinal, dictionnaire)

# Définition du nom du fichier excel en sortie et de son emplacement
date = datetime.datetime.now().strftime("%Y%m%d") # Prend la date du jour 
ExcelName = OutputFile + "_" + date + ".xlsx" # Construit le nom de fichier de sortie
ExcelFile = os.path.join(dirExcel,ExcelName) # Join la path et le nom de fichier

writer = pd.ExcelWriter(ExcelFile) # creating excel writer object
table.to_excel(writer) # write dataframe to excel
writer.close() # close the excel

print("Excel File is successfully created") # Print excel created