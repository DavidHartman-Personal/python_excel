# Excel Python utilities

This repository is based on a default Python project template.  The various resources are described below as well as actions that should be completed once the project is created.

## docs/

Contains files that create and manage documentation for the project.  These resources can utilize tools that automatically generate documentation based on following standards.  E.g. When using google doc strings in Python files, sphinx can be used to automatically create documentation that pulls code comments to create code documentation.  Setup of Sphinx, etc. has to be done in order to generate documentation.

## sample/

This is the main code directory where the main python project code resides.  This folder is really just for reference and should be removed and/or renamed based on this projects purpose.  

* [ ] Create new project folder and remove the sample/ folder

## tests/

This contains files/resources for running automated tests (unit, integration, etc.).  It could be left alone initially until real testing requirements have been figured out.

* [ ] Create a simple unit test that references actual project code.  Doing this as early as possible makes it easier to just expand these automated tests at a later time.

Makefile - The Makefile will install any required modules (e.g. Sphinx tools, etc.)
setup.py - This file is a basic python file for creating a module/library ?!? 
requirements.txt - This file will include all the python modules used/needed for this project. 

* [ ] Update setup.py with actual project information. 
* [ ] Include a diagram that shows how these objects relate, etc.
* [ ] Add utilities, etc. to create Documentation (Sphinx, etc.)
* [ ] Add decorators to classes.
* [ ] Add sphinx modules/etc. and run batch file to update docs/ folder
* [ ] Review Python Documentation repository and add general directions to this readme for generating/updating code documentation.  This can be done once excel related python code has been added.




| Object          | Description                                                     | Notes |
|-----------------|-----------------------------------------------------------------|-------|
| excel_models    | Contains main ExcelWorkbook class definition                    |       |
| excel_utilities | Contains various utilities for working with Excel spreadsheets. |       |
