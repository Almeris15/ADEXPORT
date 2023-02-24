# Adexport
Dossier contenant des scripts afin de réaliser une exportation d'Active Directory automatiquement ou bien de comparer 2 matrice issue d'une exportation d'Active Directory.

# Prérequis : 
- Avoir Python : version 3.10.9 ou ultérieur
- Avoir des accès à un Active Directory pour le script Powershell

# __** Seul le script RunProgram.py est à exécuter **__

# Utilisation
## Utilisation de Powershell
```
.\ExportAD.ps1 'DNameDNS'
```

## Utilisation de python
Ce programme permet soit de comparer 2 matrices soit d'éditer 1 matrice  
Pour crée une matrice :
``` 
.\RunProgram.py --Edit 'ADNameDNS'
```
Pour comparer 2 matrices :
``` 
.\RunProgram.py --Compare 'ADNameDNS'
```
Le résultat quelque soit l'option choisit est un tableau à double entrée dans un fichier excel.  

# Arborescence :
- POWERSHELL
    - ExportAD.ps1
- PYTHON
    - CSVtoMatriceExcel.py
    - editExcelFile.py
    - Compare.py
- TEMP
    - Fichiers CSV
- EXCEL_FILES
- RunProgram.py
- README.md

# Description :
## POWERSHELL
### ExportAD.ps1
Script powershell qui fait un export de l'ad et sort 3 csv:
- 1 pour la liste des groupes - Format : 'ADgroups_ADnameDNS_YYYYMMDD.csv'
- 1 pour la liste des utilisateurs - Format : 'ADmembers_ADnameDNS_YYYYMMDD.csv'
- 1 pour la liste utilisateur-groupes - Format : 'ADmembership_ADnameDNS_YYYYMMDD.csv'  
Penser à passer le nom de L'AD en format DNS (Distinguished Names)

## PYTHON
### CSVtoMatriceExcel.py
Script qui transforme les csv de l'export powershell en une matrice à double entrée et l'enregistre dans un fichier excel.  
Variables : 
- Nom du fichier excel en argument

### editExcelFile.py
Script qui met en forme le fichier excel de sortie du programme CSVtoMatriceExcel.py  
Variables :
- Nom du fichier excel en argument

### Compare.py
Script qui compare 2 Matrice AD groupmembership grâce à 3 CSV par matrice.  
On obtient un fichier Excel contenant 1 matrice avec des couleur pour voir rapidemment les différences.  
Variables : 
- Nom du fichier excel en argument

## TEMP
Dossier servant à stocker les fichier csv le temps de l'éxécution du programme python.

## EXCEL_FILES

## RunProgram.py
Script python à lancer, qui lance tout les script dans l'ordre.  
Passer en argument les options selon si on veut Editer une Matrice ou Comparer 2 matrice  
Vérifier que les CSV nécessaire se trouvent dans le Dossier "TEMP"  
Variables : 
- Noms de tout les autres Scripts
- Noms des fichiers CSV

## README.md
- Descriptions du programme
- Prérequis
- Utilisation
- Arborescence
- Description des différents dossier et scripts
