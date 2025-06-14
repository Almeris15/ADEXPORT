import logging
import sys
import os
import pandas as pd
from datetime import datetime
from log_config import setup_logging


# --- variables ---
dir = os.path.dirname(__file__)         # Absolute file location
dirExcel = dir + "\\..\\ExportExcel\\"  # Excel folder location
OutputFile = sys.argv[1]                # Excel file name in script call argv
dirCSV = dir + "\\..\\ImportCSV\\"      # CSV folder location
# --- End of variables ---


# --- Start of functions ---
def FileDate(fileName):
    """Function to choose file by date."""
    latest_file_name = None
    latest_file_date = 0
    file_date_str = fileName.split("_")[-1].split(".")[0]  # Assumes the format is "ADmembers_'Ad.Name'_YYYYMMDD.csv" and keeps only the date
    file_date = int(file_date_str)
    if latest_file_date is None or file_date > latest_file_date:  # replace file with the latest one
        latest_file_name = os.path.join(dirCSV, fileName)
        latest_file_date = file_date
    return latest_file_name  # returns the most recent file


def Call_files():
    """Function for file call."""
    for file_name in os.listdir(dirCSV):
        if file_name.startswith("ADmembers_") and file_name.endswith(".csv"):  # only chooses csv files starting with Admembers_
            User_file = FileDate(file_name)
        if file_name.startswith("ADgroups_") and file_name.endswith(".csv"):  # only chooses csv files starting with Adroups_
            Group_file = FileDate(file_name)
        if file_name.startswith("ADmembership_") and file_name.endswith(".csv"):  # only chooses csv files starting with Admembership_
            Membership_file = FileDate(file_name)
    return User_file, Group_file, Membership_file  # return the 3 necessary csv files


def File_to_list(filecsv):
    """Function for the first 2 csv files: User and Groups."""
    with open(filecsv, 'r') as fileR:
        List = fileR.readlines()[1:]  # read all lines except the first one
    ListFinal = [i.replace('"', '').replace("\n", '') for i in List]  # Structure the code by removing quotes and \n
    return ListFinal


def Group_Membership(filecsv):
    """Function for 3rd csv file with groupmembership."""
    with open(filecsv, 'r') as fileR:
        lineR = fileR.readlines()[1:]  # read all lines except the first one
    dico = {}
    for line in lineR:
        name, group = map(str.strip, line.strip().split(","))  # Split the csv list into 2 columns at ","
        name = name.strip('"')
        group = group.strip('"')
        dico.setdefault(group, []).append(name)  # append the new number to the existing array else create a new array
    return dico


def Generate_table(list1, list2, dico):
    """Function to generate the table."""
    tableau = []
    for name in list2:  # for each user
        row = ["X" if name in dico.get(i, []) else " " for i in list1]  # For each line put X if membership otherwise put nothing
        tableau.append(row)
    table = pd.DataFrame(tableau, columns=list1, index=list2)
    return table
# --- End of functions ---


# --- Call to functions ---
setup_logging()
User, Groups, GroupMembership = Call_files()

ListUserFinal = File_to_list(User)
ListGroupFinal = File_to_list(Groups)
dictionnaire = Group_Membership(GroupMembership)
table = Generate_table(ListGroupFinal, ListUserFinal, dictionnaire)

# Defining the output Excel file name and location
date = datetime.now().strftime("%Y%m%d")  # Take today's date
ExcelName = OutputFile + "_" + date + ".xlsx"  # Constructs the output filename
ExcelFile = os.path.join(dirExcel, ExcelName)  # Join the path and file name

writer = pd.ExcelWriter(ExcelFile)  # creating excel writer object
table.to_excel(writer)  # write dataframe to excel
writer.close()  # close the excel

print("Excel File is successfully created")
logging.info('Created File Excel : %s', ExcelName)
