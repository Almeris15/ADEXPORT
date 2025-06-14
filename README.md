# ADEXPORT
[![fr](https://img.shields.io/badge/lang-fr-green.svg)](https://github.com/Almeris15/ADEXPORT/blob/master/README.fr.md)

Program to help for reporting AD member, membership and Groups
A program containing scripts to automatically export Active Directory and either create a matrix or compare two matrices from an Active Directory export.
To create/compare these matrices, you must first export the Active Directory to obtain the CSV files. The CSV export is also included in the program.

---

## Prerequisites:
- Access to an Active Directory for the PowerShell script
- Python: version 3.10.9 or later
- Know the AD name in Distinguished Name --> ADNameDNS

---

## Usage

__** Only the RunProgram.py script needs to be run **__

### PowerShell Usage
```
.\ExportAD.ps1 'ADNameDNS'
```

---

### Python Usage
This program allows you to either create a matrix or compare two matrices using the csv files located in the ImportCSV folder. To create a matrix from an AD:
```
\\path\\python.exe \\path\\RunProgram.py --Create 'ADNameDNS'
```
To compare two matrices from the same AD:
```
\\path\\python.exe \\path\\RunProgram.py --Compare 'ADNameDNS'
```
CSV files from the same AD are compared by reverse chronological date.
The number of Excel files to compare is therefore equal to the number of dates - 1.
Regardless of the option chosen, the result is a double-entry table stored in an Excel file and stored in the ExportExcel folder.

---

## 📁File Tree:

    ADEXPORT  
    ├── ExportExcel/  
        └── Output matrix files in .xlsx  
    ├── ImportCSV/  
        └── CSV files  
    ├── POWERSHELL/  
        └──ExportAD.ps1  
    ├── PYTHON/  
        ├── CSVtoMatriceExcel.py  
        ├── Compare.py  
        ├── editExcelFile.py  
        └── log_config.py    
    ├── ExportAD.log  
    ├── LICENSE  
    ├── README_fr.md  
    ├── README_en.md  
    └── RunProgram.py

If the ImportCSV and ExportExcel folders do not exist, please create them according to this tree structure and the case of their respective names.

---

## Tree structure description
### POWERSHELL folder
#### ExportAD.ps1
PowerShell script that exports the Active Directory and outputs 3 CSV files:
- 1 for the list of groups - Format: 'ADgroups_ADnameDNS_YYYYMMDD.csv'
- 1 for the list of users - Format: 'ADmembers_ADnameDNS_YYYYMMDD.csv'
- 1 for the user-group list - Format: 'ADmembership_ADnameDNS_YYYYMMDD.csv'
Remember to pass the AD name in DNS (Distinguished Names) format

---

### PYTHON File
#### CSVtoMatriceExcel.py
Script that transforms the CSV values ​​from the PowerShell export into a two-way matrix and saves it to an Excel file.
Variables:
- ADNameDNS as an argument

---

#### editExcelFile.py
Script that formats the Excel file output from the CSVtoMatriceExcel.py program
Variables:
- ADNameDNS as an argument

---

#### Compare.py
Script that compares two AD groupmembership matrices.
The result is an Excel file containing one matrix with colors (green and red) to quickly see the differences.
Variables:
- ADNameDNS as an argument

---

### IMPORT_CSV folder
Folder used to store the imported CSV files for program execution.

---

### EXPORT_EXCEL folder
Folder used to store the Excel files output from the program execution.

---

### RunProgram.py
Python script to run, which launches all the scripts in order.

Pass the options as arguments depending on whether you want to Edit a Matrix or Compare 2 Matrices.
Check that the necessary CSV files are in the "IMPORT_CSV" folder.
Variables:
- Script names.py
- ADNameDNS as an argument
- One of the options seen in the 'Usage' section above

---

### README.md
- Program descriptions
- Prerequisites
- Usage
- Tree structure
- Description of the different folders and scripts

---

## 📜 License

MIT License – free to use, share, and modify. Contributions welcome!

---