#NAME : Compare.py
#Author : Almeris15
#Description : 
#  
# Transformer la sortie des informations AD de liste à dict pour traitement.
# les csv sont transformes en une matrice dans un dataframe.
# les dataframes sont ensuite comparé pour avoir un résultat à T+1.
# 
# Changelog:
# 1.0.0 - Initial release

import sys
import os 
from collections import defaultdict
from datetime import datetime
import pandas as pd
from openpyxl.styles import PatternFill, Border, Side, Alignment

# ----- Variables -----
dir = os.path.dirname(__file__) # Emplacement du fichier grâce à son chemin absolue
dirExcel = dir + "\\..\\EXCEL_FILES\\"
dirTemp = dir + "\\..\\TEMP\\"
OutputFile = sys.argv[1] # Nom du fichier Excel = 1er argument en appel de script
color = 'F2F2F2' # Color for the cell in hex code
colorborder = 'FF000000' # color of the border of cells
rotation = 45 # Witch angle to rotate the text in a cell
widths = 4.3 # widths of all columns
widthsFirst = 20 # widths of the first column
# ----- Fin des variables -----

# ----- Début des fonctions -----
# Fonction pour appel de fichier
def Call_files():
    file_names_list= []
    file_groups_list = []
    file_membership_list = []
    for file_name in os.listdir(".\\TEMP\\"): # se place dans le dossier Temp
        if file_name.startswith("ADmembers_") and file_name.endswith(".csv"): # ne choisit que les fichiers csv commençant par Admembers_
            if file_name.split('_')[1] == OutputFile: # Vérifie que le ADNameDNS donné en argument de programme soit dans le nom du fichier
                file_names_list.append(file_name)
        if file_name.startswith("ADgroups_") and file_name.endswith(".csv"): # ne choisit que les fichiers csv commençant par Adroups_
            if file_name.split('_')[1] == OutputFile: # Vérifie que le ADNameDNS donné en argument de programme soit dans le nom du fichier
                file_groups_list.append(file_name)
        if file_name.startswith("ADmembership_") and file_name.endswith(".csv"): # ne choisit que les fichiers csv commençant par Admembership_
            if file_name.split('_')[1] == OutputFile: # Vérifie que le ADNameDNS donné en argument de programme soit dans le nom du fichier
                file_membership_list.append(file_name)
    file_ADnames_list_sort  = sorted(file_names_list, key=lambda x: datetime.strptime(x.split('_')[-1].split(".")[0], '%Y%m%d'), reverse=True) # Sort la liste de fichier trié par date : du plus récent au plus ancien
    file_ADgroups_list_sort = sorted(file_groups_list, key=lambda x: datetime.strptime(x.split('_')[-1].split(".")[0], '%Y%m%d'), reverse=True) # Idem que au dessus
    file_ADmembership_list_sort = sorted(file_membership_list, key=lambda x: datetime.strptime(x.split('_')[-1].split(".")[0], '%Y%m%d'), reverse=True) # Idem que au dessus
    return file_ADnames_list_sort, file_ADgroups_list_sort, file_ADmembership_list_sort # return les 3 fichier csv nécessaires

# Fonction pour avoir les 2 fichier CSV de chaque list pour comparaison
def search_file_list(file_list):
    file2 = file_list[0]
    file1 = file_list[1]
    print(file1, file2)
    file_list.pop(0)
    file_date_str = file2.split("_")[-1].split(".")[0]  # Présume que le format est "ADmembers_'Ad.Name'_YYYYMMDD.csv" et garde que la date
    file_date = int(file_date_str)
    return file1, file2, file_date

# Fonction pour les 2 premier fichier csv avec les User et les Groupes
def FileToList(filecsv):
    file_read = os.path.join(dirTemp, filecsv)
    with open(file_read, 'r') as fileR:
        List = fileR.readlines()
    ListFinal = [i.replace('"','').replace("\n",'') for i in List[1:]]
    return ListFinal

# Assemble les 2 listes
def concat_lists(listA, listB):
    listR = listA[:]
    seen = set(listR)
    listBOnly = [x for x in listB if x not in seen and not seen.add(x)]
    listAOnly = [x for x in listA if x not in listB]
    listR.extend(listBOnly)
    listR.sort()
    return listR, listBOnly, listAOnly

# Fonction pour le 3ème fichier csv avec groupemembership 
def Group_Membership(filecsv):
    file_read = os.path.join(dirTemp, filecsv)
    with open(file_read, 'r') as fileR:
        lineR = fileR.readlines()
        dico = defaultdict(list)
        for lineR in map(str.rstrip, lineR):
            Name, Group = lineR.split(",")
            Name = Name.strip('"')
            Group = Group.strip('"')
            dico[Group].append(Name)
    del dico['GroupName']
    return dict(dico)

# Fonction pour générer le tableau
def Generate_table(list1, list2, dico):
    tableau = []
    for name in list2:
        row = [name in dico.get(group, []) and 'X' or ' ' for group in list1] # Met des X si user appartient au group sinon met rien / espace
        tableau.append(row) # Crée la ligne
    table = pd.DataFrame(tableau, columns=list1, index=list2) # Crée la matrice
    return table

# Fonction pour comparé les 2 dataframes
def Compare_table(df1,df2,OutputExcel):
    # Comparaison des 2 matrices au niveau des 'X'
    mask1 = df1 == "X"
    mask2 = df2 == "X"

    writer = pd.ExcelWriter(OutputExcel, engine="openpyxl") # création d'un fichier excel de sortie
    df2.to_excel(writer, index=True) # mettre la matrice la plus récente dans le excel
    ws = writer.book.active
    
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid") # setup couleur rouge
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid") # setup couleur verte
    
    for row_index, row in enumerate(mask1.values):
        for col_index, cell in enumerate(row):
            current_cell = ws.cell(row=row_index + 2, column=col_index + 2)
            current_cell.alignment = Alignment(horizontal='center', vertical='center')
            current_cell.border = Border(top=Side(border_style='thin', color=colorborder), bottom=Side(border_style='thin', color=colorborder)) # Set border
            if row_index % 2 != 0:
                current_cell.fill = PatternFill(patternType='solid', fgColor=color) # Set color
            if cell and not mask2.iloc[row_index, col_index]:
                current_cell.fill = red_fill # Set color
            elif not cell and mask2.iloc[row_index, col_index]:
                current_cell.fill = green_fill # Set color
                
    for column in ws.columns:
        ws.column_dimensions[column[0].column_letter].width = widths
        
    ws.column_dimensions['A'].width = widthsFirst
    first_row = ws[1]
    first_column = ws['A']
    
    for cell in first_row:
        cell.alignment = Alignment(textRotation=rotation) # rotation of the text for the group name
        cell.border = Border(right=Side(border_style='thin', color=colorborder),left=Side(border_style='thin', color=colorborder),bottom=Side(border_style='thin', color=colorborder)) # Set border
        if cell.value in listGroupGreen:
            cell.fill = green_fill # Set color
        if cell.value in listGroupRed:
            cell.fill = red_fill # Set color
    
    for cell in first_column:
        cell.alignment = Alignment(horizontal='left')
        cell.border = Border(right=Side(border_style='thin', color=colorborder), top=Side(border_style='thin', color=colorborder), bottom=Side(border_style='thin', color=colorborder)) # Set border
        if cell.value in listUserGreen:
            cell.fill = green_fill # Set color
        if cell.value in listUserRed:
            cell.fill = red_fill # Set color
    
    writer.close()
# ----- Fin des fonctions -----

# Appels aux fonctions
User_file_list, Groups_file_list, GroupMembership_file_list = Call_files()

while len(User_file_list) > 1 : # Quand il y a + qu'un fichier dans la liste de fichier alors comparaison
    # retourne les fichiers csv nécessaires pour la comparaison    
    User_file1, User_file2, user_file_date = search_file_list(User_file_list)
    Group_file1, Group_file2, group_file_date = search_file_list(Groups_file_list)
    Membership_file1, Membership_file2, membership_file_date = search_file_list(GroupMembership_file_list)

    if user_file_date == group_file_date == membership_file_date: # Condition pour s'assurer que les fichier ont la même date
        ListUser1 = FileToList(User_file1)
        ListUser2 = FileToList(User_file2)
        ListGroup1 = FileToList(Group_file1)
        ListGroup2 = FileToList(Group_file2)
        dict1 = Group_Membership(Membership_file1)
        dict2 = Group_Membership(Membership_file2)

        ListUserFinal, listUserGreen, listUserRed = concat_lists(ListUser1, ListUser2)
        ListGroupFinal, listGroupGreen, listGroupRed = concat_lists(ListGroup1, ListGroup2)

        table1 = Generate_table(ListGroupFinal, ListUserFinal, dict1)
        table2 = Generate_table(ListGroupFinal, ListUserFinal, dict2)

        date = str(user_file_date)
        fileName = OutputFile + "_" + date + ".xlsx"
        Excelfile = os.path.join(dirExcel,fileName)

        tableFinal = Compare_table(table1,table2,Excelfile)
        print(f'Create "{fileName}"') # print the file edit
    else:
        print("!! Erreur au niveau des dates des fichiers csv !!")

print("Compare.py Finish") # print end of program
