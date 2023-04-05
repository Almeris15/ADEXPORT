# ADEXPORT
Programme contenant des scripts afin de réaliser une exportation d'Active Directory automatiquement et soit crée une matrice soit de comparer 2 matrices issue d'une exportation d'Active Directory.  
Pour crée/comparé ces matrices il faut dans un premier temps faire un export de l'active directory afin d'obtenir les fichier CSV. L'export des csv est aussi inclu dans le programme.

# Prérequis : 
- Avoir des accès à un Active Directory pour le script Powershell
- Avoir Python : version 3.10.9 ou ultérieur
- Connaître le nom de l'AD en Distinguished Name  --> ADNameDNS

# __** Seul le script RunProgram.py est à exécuter **__

# Utilisation
## Utilisation Powershell
```
.\ExportAD.ps1 'ADNameDNS'
```
## Utilisation Python
Ce program permet soit de crée 1 matrice soit de comparer 2 matrices grâce aux csv se trouvant dans le dossier ImportCSV.  
Pour crée une matrice d'un AD :
``` 
\\path\\python.exe \\path\\RunProgram.py --Create 'ADNameDNS'
```
Pour comparer 2 matrices d'un même AD :
``` 
\\path\\python.exe \\path\\RunProgram.py --Compare 'ADNameDNS'
```
Les fichiers CSV d'un même AD sont comparé par date antichronologique  
Le nombre de fichier Excel de comparaison est donc égal aux nombres de date - 1  
Quelque soit l'option choisit le résultat est un tableau à double entrée stocké dans un fichier excel est entreposé dans le dossier ExportExcel.  

# Arborescence :
- EXPORT_EXCEL
    - Fichiers de matrice de sortie en .xlsx
- IMPORT_CSV
    - Fichiers CSV
- POWERSHELL
    - ExportAD.ps1
- PYTHON
    - CSVtoMatriceExcel.py
    - Compare.py
    - editExcelFile.py
    - log_config.py
- ExportAD.log
- LICENSE
- README.md
- RunProgram.py

Si les dossiers ImportCSV et ExportExcel n'existent pas merci de les crées suivant cette arborescence et la casse dde leurs noms respectifs  

# Description de l'arborescence
## Dossier POWERSHELL
### ExportAD.ps1
Script powershell qui fait un export de l'ad et sort 3 csv:
- 1 pour la liste des groupes - Format : 'ADgroups_ADnameDNS_YYYYMMDD.csv'
- 1 pour la liste des utilisateurs - Format : 'ADmembers_ADnameDNS_YYYYMMDD.csv'
- 1 pour la liste utilisateur-groupes - Format : 'ADmembership_ADnameDNS_YYYYMMDD.csv'  
Penser à passer le nom de L'AD en format DNS (Distinguished Names)

## Dossier PYTHON
### CSVtoMatriceExcel.py
Script qui transforme les csv de l'export powershell en une matrice à double entrée et l'enregistre dans un fichier excel.  
Variables : 
- ADNameDNS en argument

### editExcelFile.py
Script qui met en forme le fichier excel de sortie du programme CSVtoMatriceExcel.py  
Variables :
- ADNameDNS en argument

### Compare.py
Script qui compare 2 matrices AD groupmembership.  
On obtient un fichier Excel contenant 1 matrice avec des couleur (vert et rouge) afin de voir rapidemment les différences.  
Variables : 
- ADNameDNS en argument

## Dossier IMPORT_CSV
Dossier servant à stocker les fichier csv en import pour l'éxécution du programme.

## Dossier EXPORT_EXCEL
Dossier servant à stocker les fichier excel en sortie d'éxécution du programme.

## RunProgram.py
Script python à lancer, qui lance tout les script dans l'ordre.  
Passer en argument les options selon si on veut Editer une Matrice ou Comparer 2 matrice  
Vérifier que les CSV nécessaire se trouvent dans le Dossier "IMPORT_CSV "  
Variables : 
- Noms des Scripts .py
- ADNameDNS en argument
- une des Option vu dans la partie 'Utilisation' au dessus

## README.md
- Descriptions du programme
- Prérequis
- Utilisation
- Arborescence
- Description des différents dossier et scripts
