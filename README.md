# EDA-Seminar-Organization-Grading

This repository contains a python project to manage seminars in the chair of Design Automation.

# Project Structure

```
eda-seminar-organization-grading/
├── DataSources (Contains template xlsx sheets to prepare final sheets)
│   ├── Foik_GradingSheetSeminar.xlsx (template for supervisor grading sheets [Script2])
│   ├── input_updated.xlsx (test file with new students from TUMonline[Script1])
│   ├── input.xlsx (Initial test file from TUMonline[Script1])
│   ├── master_sheet_manually_updated.xlsx (Manually updated project details in Master Sheet  generated by Script1)
│   └── master_sheet.xlsx (template sheet for generating Master Sheet from Script1)
├── Documentation
│   └── generate_directory_tree.py (Script to generated Directory structure of the project)
│   └── Project-structure-graphical.pdf (Project structure slides)
│   └── Project-structure-graphical.pptx (Project structure slides)
├── html (This Directory contains documentation of Scripts in HTML format)
│   └── ConfirmedStudents.html (pdoc for Script1)
│   └── GraderSheets.html (pdoc for Script2)
├── LICENSE 
├── Makefile
├── MANIFEST.in
├── OutputFiles (Output files/sheets generated by Scripts)
│   ├── GraderSheets (Grading Sheet for supervisor generated by Sheet2)
│   │   ├── Bing Li_GradingSheetSeminar.xlsx
│   │   ├── Foik_GradingSheetSeminar.xlsx
│   │   ├── Mengchu Li_GradingSheetSeminar.xlsx
│   │   ├── Mettler_GradingSheetSeminar.xlsx
│   │   ├── Moradi_GradingSheetSeminar.xlsx
│   │   ├── Müller-Gritschneder_GradingSheetSeminar.xlsx
│   │   ├── Neuner_GradingSheetSeminar.xlsx
│   │   └── Stahl_GradingSheetSeminar.xlsx
│   ├── master_sheet_2021_01_07_12_44_45.xlsx (Master Sheets generated by Script1 with timestamp)
│   └── master_sheet_2021_01_07_12_46_52.xlsx (do)
├── README.md
├── requirements.txt (follow the README.md to use this file to install python3 dependencies)
├── setup.py (project setup file)
├── tests (test files)
│   └── test_add.py
└── toolscripts (Here are the source scripts that produce output sheets)
    ├── ConfirmedStudents.py (aka Script1)
    ├── GraderSheets.py (aka Script2)
    ├── __init__.py
    └── __pycache__
        └── ConfirmedStudents.cpython-38.pyc

```

# Development Section

## Create and activate environment
```python3 -m venv ./venv```  
```source venv/bin/activate```

## Install packages
```pip install -r requirements.txt```

## Run Script1
* ```python toolscripts/ConfirmedStudents.py DataSources/input.xlsx```

* For HauptSeminar use the flag -HS
```python toolscripts/ConfirmedStudents.py DataSources/input.xlsx -HS```

* for update with an existing Master file
```python toolscripts/ConfirmedStudents.py <new_file_from_TUMonline> --update=<manually_updated_file> ```

for example:
```python toolscripts/ConfirmedStudents.py DataSources/input_updated.xlsx --update=OutputFiles/master_sheet_2020_12_10_17_11_00.xlsx```

It is to be noted that when updating a manually updated file with latest registered students, -HS key is **not** required. (use -HS for only **first** time in case of Hauptseminar, because it does add extra columns and shuffle the review pattern).
Only ```--update=<manually updated file>``` is sufficient.

## Run Script2
* ```python toolscripts/GraderSheets.py DataSources/master_sheet_manually_updated.xlsx```

Grader Xlsx files are generated in this path :
```OutputFiles/GraderSheets/```

## Generate pdoc documentation in the directory html/
```pdoc --html toolscripts/ConfirmedStudents.py```

To overwrite
```pdoc --html toolscripts/ConfirmedStudents.py --force``` 

## Changelog(17.12.2020)
* used openpyxl
* -HS switch is working 
* replaced headers while xlsx-->dataframe
* filter out students who does not have fixed place
* add columns for project title and mentor
* if Hauptseminar added two extra columns
* writeback dataframe to xlsx masterfile in /OutputFiles 
* added pycharm configs to run test scripts

==========

* Added shuffling to review to and review for
* update the master file with a new input file(with extra student entry)
* added help feature for script(argparse)

## Script2 (Ready to test)
* Generated grader files for individual supervisor with respective names in the master sheet.
* Added overview page with cell customized with grader name
* Added Paper Grading Sheet with multi student support
* Added Review Grading Sheet with multi student support
* Added Presentations Sheet for all talks in chronological order

## Changelog(24.01.2021)
* Added pdoc3 package for documentation in requirements
* Added Project directory structure in README
* Cleaned and documented Script1 and Script2
* Documentation for the whole project in Graphical presentation(Block Diagram)
* pdoc html files for both scripts
