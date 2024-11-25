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
import coloredlogs

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

coloredlogs.install(level=logging.INFO,
                    fmt="%(asctime)s %(hostname)s %(name)s %(filename)s line-%(lineno)d %(levelname)s - %(message)s",
                    datefmt='%H:%M:%S')

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

    def get_table_data(self, table_data):
        """ Get a Excel table from a worksheet and return it as a dictionary object.

        Args:
            table_data:

        """
        return_dictionary_list = []
        worksheet_table = []
        # Grab the 'data' from the table
        rows_list = []
        for row in table_data:
            # Get a list of all columns in each row
            cols = []
            for col in row:
                cols.append(col.value)
                rows_list.append(cols)
        header_columns = [col.upper() for col in rows_list[0]]
        for data_row in rows_list[1:]:
            row_dictionary = zip(header_columns, data_row)
            return_dictionary_list.append(dict(row_dictionary))
        logging.info("Table data dictionary [%s]:", str(return_dictionary_list))
        return return_dictionary_list

    def get_worksheets(self) -> []:
        """Returns an array of the worksheet objects in the workbook

        Returns:
            array object:
        """
        return_worksheets_list = []
        for ws in self.workbook.worksheets:
            return_worksheets_list.append(ws.title)
            # logging.debug("Looking at worksheet: [%s]", str(ws))
            # let's add it to the dictionary of worksheets if we haven't already
            if not self.worksheets.get(ws.title):
                worksheet_object = {}
                worksheet_title = ws.title
                worksheet_tables = {}
                for tbl in ws._tables:
                    ws_table_name = tbl.name
                    table_dictionary = self.get_table_data(ws[tbl.ref])
                    worksheet_tables['TABLE_NAME'] = table_dictionary
                worksheet_object['WORKSHEET_NAME'] = worksheet_title
                self.worksheets[ws.title] = worksheet_object
        # logging.debug("Worksheets [%s]:", str(return_worksheets_list))
        return return_worksheets_list

