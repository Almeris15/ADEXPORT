import argparse
import subprocess
import os
import datetime
import sys
import logging
from PYTHON.log_config import setup_logging

# --- Start of variables ---
dir = os.path.dirname(__file__)  # Import absolute path
# Path for scripts and programs
program1 = os.path.join(dir, ".\\PYTHON\\CSVtoMatriceExcel.py")
program2 = os.path.join(dir, ".\\PYTHON\\editExcelFile.py")
program3 = os.path.join(dir, ".\\PYTHON\\Compare.py")
# --- End of variables ---


# --- Start of functions ---
def RunProgram(program):
    """ Function to launch the Python scripts."""
    try:
        print("Program run: ", program)
        logging.info('launching the script: %s', program)
        subprocess.run(["python", program, output_file_name])  # Lancement du script avec options
    except FileNotFoundError:
        print(f"Script {program} not found")
        logging.error('Script %s not found', program)


def DeleteFile(csv):
    """ Function to delete the csv."""
    try:
        os.remove(csv)
        print("supression de ", csv)
    except FileNotFoundError:
        pass


def install(package):
    """ Functions to install the missing necessary packages."""
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
# --- End of functions ---


# --- Call for functions ---
# Check and/or install pkg
install('pandas')
install('openpyxl')

setup_logging()

# Initialize the argument parser with option
parser = argparse.ArgumentParser(description='Process csv into Excel matrix')  
parser.add_argument("--Create", help="Create a matrix using csv", type=str, nargs=1)
parser.add_argument("--Compare", help="Compare 2 matrix using csv", type=str, nargs=1)

try:
    args = parser.parse_args()  # Parse the command-line arguments
except Exception as e:
    logging.error("Error command")

date = datetime.datetime.now().strftime("%Y%m%d")  # Get date with the form YYYYMMDD

if args.Create:
    output_file_name = args.Create[0]
    print(f"Creating a matrix. Output file name:'{output_file_name}_{date}.xlsx'")
    logging.info("Creation of a matrix. Name of the file in exit:'%s_%s.xlsx'", output_file_name, date)
    RunProgram(program1)  # Launch of the CSVToMatrice program
    RunProgram(program2)  # Launch of the EditExcel program
elif args.Compare:
    output_file_name = args.Compare[0]
    print(f"Comparison of 2 matrix.")
    logging.info("Compare 2 matrix")
    RunProgram(program3)  # Launch the python program to compare 2 matrices

    print("Please choose an option")
    logging.error("Error no option has been chosen")

print("End the Program")
logging.info('End of program')
