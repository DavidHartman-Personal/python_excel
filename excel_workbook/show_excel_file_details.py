"""Read and Show Excel workbook details

This script will open an Excel Workbook and then show some details about the workbook.  It uses configuration files, and references the data/ directory.

"""

import argparse
import os
import excel_workbook
import logging
import coloredlogs

#: This effectively defines the root of the project and so adding ..\, etc. is not needed in config files
PROJECT_ROOT_DIR = os.path.dirname(os.path.dirname(__file__))

#: Directory that contains configuration files
CONF_DIR = os.path.join(PROJECT_ROOT_DIR, 'conf')

#: Directory were data files/extracts/reports will be stored
DATA_DIR = os.path.join(PROJECT_ROOT_DIR, 'data')

#: Directory were data files/extracts/reports will be stored
EXCEL_FILE_DIR = os.path.join(PROJECT_ROOT_DIR, 'excel_files')

#: Configuration file path.  Uses environment variable if none is defined.
FILENAME_INPUT_CONFIG = os.environ.get('CONFIG_FILE_PATH',
                                       os.path.join(CONF_DIR, 'excel.conf'))

coloredlogs.install(level=logging.DEBUG,
                    fmt="%(asctime)s %(hostname)s %(name)s %(filename)s line-%(lineno)d %(levelname)s - %(message)s",
                    datefmt='%H:%M:%S')

def main():
    """The main method for this script.

    """

    source_spreadsheet_name = "sample.xlsx"
    source_spreadsheet_file = os.path.join(EXCEL_FILE_DIR, source_spreadsheet_name)

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        'input_file',
        type=str,
        help="The spreadsheet file to print the columns of",
        nargs='?',
        default=source_spreadsheet_file
    )

    logging.info("Getting excel info")
    args = parser.parse_args()
    excel_wb = excel_workbook.ExcelWorkbook(source_spreadsheet_file)
    excel_ws = excel_wb.get_worksheets()

    # get_spreadsheet_cols(args.input_file, print_cols=True)


if __name__ == "__main__":
    main()
