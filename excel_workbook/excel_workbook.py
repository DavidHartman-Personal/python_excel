"""This is the main Excel Workbook class

Creates an Excel workbook class object.

    Attributes:
            workbook_filename (str): The full filename path to the Excel workbook
            worksheets (int):  A dictionary containing the Excel worksheets in the Excel workbook
            defined_names (str): An array of the defined names in the Excel workbook


    Methods
    -------

"""
__version__ = '2024.11'
__author__ = 'David Hartman'

import logging
import os
import sys
from openpyxl import load_workbook

#: This effectively defines the root of this project and so adding ..\, etc is not needed in config files
PROJECT_ROOT_DIR = os.path.dirname(os.path.dirname(__file__))

# Add script directory to the path to allow searching for modules
sys.path.insert(0, PROJECT_ROOT_DIR)

#: Directory that contains configuration files
CONF_DIR = os.path.join(PROJECT_ROOT_DIR, 'conf')

tab_separator = "\t"
comma_separator = ", "

# import any functions inside or outside of this module.  If no helpers are needed it can be removed from here
# as well as remove the file from the excel_workbook module.
# from . import helpers

class ExcelWorkbook:

    def __init__(self,
                 workbook_filename):
        self.workbook_filename = workbook_filename
        self.worksheets = {}
        self.defined_names = []
        if not os.path.exists(workbook_filename):
            logging.error("File not found [%s], if writing pass write_flag = True", str(workbook_filename))
            self.workbook = None
        else:
            self.workbook = load_workbook(filename=workbook_filename, data_only=True)
            logging.info("Creating Excel object based on file [%s]", str(workbook_filename))
