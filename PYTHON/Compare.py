import logging
import sys
import os
import pandas as pd
from datetime import datetime
from collections import defaultdict
from log_config import setup_logging
from openpyxl.styles import PatternFill, Border, Side, Alignment


# ----- Variables -----
dir = os.path.dirname(__file__)     # File location using its absolute path
dirExcel = dir + "\\..\\ExportExcel\\"
dirTemp = dir + "\\..\\ImportCSV\\"
OutputFile = sys.argv[1]            # Excel file name = 1st argument in script call
color = 'F2F2F2'                    # Color for the cell in hex code
colorborder = 'FF000000'            # Color of the border of cells
rotation = 45                       # Which angle to rotate the text in a cell
widths = 4.3                        # Widths of all column
widthsFirst = 20                    # widths of the first column
# ----- End of variables -----


# ----- Start of functions -----
def Call_files():
    """Function for file call."""
    file_names_list = []
    file_groups_list = []
    file_membership_list = []
    for file_name in os.listdir(dirTemp): 
        if file_name.startswith("ADmembers_") and file_name.endswith(".csv"):  # only chooses csv files starting with Admembers_
            if file_name.split('_')[1] == OutputFile:  # Checks that the ADNameDNS given as a program argument is in the file name
                file_names_list.append(file_name)
        if file_name.startswith("ADgroups_") and file_name.endswith(".csv"):  # only chooses csv files starting with Adroups_
            if file_name.split('_')[1] == OutputFile:
                file_groups_list.append(file_name)
        if file_name.startswith("ADmembership_") and file_name.endswith(".csv"):  # only chooses csv files starting with Admembership_
            if file_name.split('_')[1] == OutputFile:
                file_membership_list.append(file_name)
    file_ADnames_list_sort = sorted(file_names_list,
                                    key=lambda x: datetime.strptime(x.split('_')[-1].split(".")[0], '%Y%m%d'),
                                    reverse=True)  # Outputs the file list sorted by date: from newest to oldest
    file_ADgroups_list_sort = sorted(file_groups_list,
                                    key=lambda x: datetime.strptime(x.split('_')[-1].split(".")[0], '%Y%m%d'),
                                    reverse=True)  
    file_ADmembership_list_sort = sorted(file_membership_list,
                                        key=lambda x: datetime.strptime(x.split('_')[-1].split(".")[0], '%Y%m%d'),
                                        reverse=True)  
    return file_ADnames_list_sort, file_ADgroups_list_sort, file_ADmembership_list_sort  # return the 3 necessary csv files


def search_file_list(file_list):
    """function to have the 2 CSV files of each list for comparison."""
    file2 = file_list[0]
    file1 = file_list[1]
    print(file1, file2)  # print each compared file to the console
    logging.info('Compare %s et %s', file1, file2) # print in the log file
    file_list.pop(0)
    file_date_str = file2.split("_")[-1].split(".")[0]  # Assumes the format is "ADmembers_'Ad.Name'_YYYYMMDD.csv" and keeps only the date
    file_date = int(file_date_str)
    return file1, file2, file_date


def FileToList(filecsv):
    """function for the first 2 csv files with Users and Groups."""
    file_read = os.path.join(dirTemp, filecsv)  # Absolute path of csv file location
    with open(file_read, 'r') as fileR:
        List = fileR.readlines()
    ListFinal = [i.replace('"', '').replace("\n", '') for i in List[1:]]  # Structure the code by removing the quotes and \n remove the 1st line
    return ListFinal


def concat_lists(listA, listB):
    """Combine the 2 lists."""
    listR = listA[:]  # takes the value listA
    seen = set(listR)
    listBOnly = [x for x in listB if x not in seen and not seen.add(x)]  # if in listB but not in listA
    listAOnly = [x for x in listA if x not in listB]  # if in listA but not in listB
    listR.extend(listBOnly)  # added listBOnly to list A to have the whole list with only 1 occurrence
    listR.sort()
    return listR, listBOnly, listAOnly


def Group_Membership(filecsv):
    """function for the 3rd csv file with groupmembership."""
    file_read = os.path.join(dirTemp, filecsv)  # Absolute path of csv file location
    with open(file_read, 'r') as fileR:
        lineR = fileR.readlines()
        dico = defaultdict(list)  # creating a default dict, subtype of dict
        for lineR in map(str.rstrip, lineR):
            Name, Group = lineR.split(",")  # Split the csv list into 2 columns at ","
            Name = Name.strip('"')  # Remove the quotes
            Group = Group.strip('"')
            dico[Group].append(Name)
    del dico['GroupName']
    return dict(dico)


def Generate_table(list1, list2, dico):
    """function to generate a table using csv."""
    tableau = []
    for name in list2:
        row = [name in dico.get(group, []) and 'X' or ' ' for group in list1]  # Put X if user belongs to the group otherwise put nothing/space
        tableau.append(row)
    table = pd.DataFrame(tableau, columns=list1, index=list2)  # Create the matrix
    return table


def bordure_cell(cell, rightborder, leftborder, topborder, bottomborder, colors):
    """Set border of cell."""
    cell.border = Border(right=Side(border_style=rightborder, color=colors),
                        left=Side(border_style=leftborder, color=colors),
                        top=Side(border_style=topborder, color=colors),
                        bottom=Side(border_style=bottomborder, color=colors))


def Compare_table(df1, df2, OutputExcel):
    """function to compare the 2 dataframes."""
    mask1 = df1 == "X"  # takes the position of all Xs in matrix 1
    mask2 = df2 == "X"  # takes the position of all Xs in matrix 2

    writer = pd.ExcelWriter(OutputExcel, engine="openpyxl")  # creating an output excel file
    df2.to_excel(writer, index=True)  # put the most recent matrix in the excel
    ws = writer.book.active

    # --- start of compare_table variables ---
    red_fill = PatternFill(start_color="FF0000",
                            end_color="FF0000",
                            fill_type="solid")
    green_fill = PatternFill(start_color="00FF00",
                             end_color="00FF00",
                             fill_type="solid")
    list_column_green = []
    list_column_red = []
    list_row_green = []
    list_row_red = []
    # --- End of compare_table variables ---


    # --- start of compare functions ---
    def fill_column(list_column, fill):
        """function to fill the columns."""
        if list_column != []:
            for col_num in list_column:
                # Takes each cell in each row corresponding to the column number
                for row in ws.iter_rows(min_row=2, min_col=col_num, max_col=col_num):
                    for cell in row:
                        cell.fill = fill


    def fill_row(list_row, fill):
        """function to fill the lines."""
        if list_row != []:
            for row_num in list_row:
                for cell in ws[row_num][1:]:  # takes each cell in the row
                    cell.fill = fill


    def fill_column_with_color(cell_axe, listcolor, listcoloraxe, color):
        """function to color the columns defined to fill."""
        for cell in ws[1]:
            if cell.value in listcolor:
                cell.fill = color
                listcoloraxe.append(cell_axe)


    def fill_row_with_color(cell_axe, listcolor, listcoloraxe, color):
        """function to color the lines defined to be filled."""
        for cell in ws['A']:
            if cell.value in listcolor:
                cell.fill = color
                listcoloraxe.append(cell_axe)
    # --- end of compare functions ---


    for row_index, row in enumerate(mask1.values):  # take mask 1 with the position of all the X not placed in the excel
        for col_index, cell in enumerate(row):
            current_cell = ws.cell(row=row_index + 2, column=col_index + 2)
            current_cell.alignment = Alignment(horizontal='center', vertical='center')  # centre le texte
            bordure_cell(current_cell, None, None, 'thin', 'thin', colorborder)
            if row_index % 2 != 0:
                # excel layout for all even rows
                current_cell.fill = PatternFill(patternType='solid', fgColor=color)
            if cell and not mask2.iloc[row_index, col_index]:
                # compares the position of the X between the 2 matrices, condition if in the 1st matrix and not in the 2nd
                current_cell.fill = red_fill
            elif not cell and mask2.iloc[row_index, col_index]:
                # if in the 2nd matrix and not in the 1st
                current_cell.fill = green_fill

    for column in ws.columns:
        ws.column_dimensions[column[0].column_letter].width = widths  # set the width of each column
    ws.column_dimensions['A'].width = widthsFirst  # set the width of the 1st column

    for cell in ws[1]:  # Take the 1st line
        cell.alignment = Alignment(textRotation=rotation)  # rotation of the text for the group name
        bordure_cell(cell, 'thin', 'thin', None, 'thin', colorborder)  # Set border

    for cell in ws['A']:  # Take the 1st column
        cell.alignment = Alignment(horizontal='left')
        bordure_cell(cell, 'thin', None, 'thin', 'thin', colorborder)  # Set border

    fill_column_with_color(cell.column, listUserGreen, list_column_green, green_fill)
    fill_column_with_color(cell.column, listUserRed, list_column_red, red_fill)
    fill_row_with_color(cell.row, listUserGreen, list_row_green, green_fill)
    fill_row_with_color(cell.row, listUserRed, list_row_red, red_fill)
    fill_column(list_column_green, green_fill)
    fill_column(list_column_red, red_fill)
    fill_row(list_row_green, green_fill)
    fill_row(list_row_red, red_fill)

    writer.close()
# ----- End of functions -----


# ----- Calls to functions -----
setup_logging()

User_file_list, Groups_file_list, GroupMembership_file_list = Call_files()

while len(User_file_list) > 1:  # When there is more than one file in the file list, compare
# returns the CSV files needed for the comparison
    User_file1, User_file2, user_file_date = search_file_list(User_file_list)
    Group_file1, Group_file2, group_file_date = search_file_list(Groups_file_list)
    Membership_file1, Membership_file2, membership_file_date = search_file_list(GroupMembership_file_list)

    if user_file_date == group_file_date == membership_file_date:  # Condition to ensure that the files have the same date
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
        Excelfile = os.path.join(dirExcel, fileName)

        tableFinal = Compare_table(table1, table2, Excelfile)
        print(f'Create "{fileName}"')  # print the file edit
        logging.info('Created file %s',fileName)
    else:
        print("!! Error in csv file dates !!")
        logging.error("!! Error in csv file dates!!")


print("Compare.py Finish")
logging.info('End Compare.py')
