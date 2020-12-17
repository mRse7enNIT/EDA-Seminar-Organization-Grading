# EDA-Seminar-Organization-Grading

This repository contains a python project to manage seminars in the chair of Design Automation.

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

## Script2 (On progress)
* Generated grader files for individual supervisor with respective names in the master sheet.
* Added overview page with cell customized with grader name
