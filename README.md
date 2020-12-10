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
```python toolscripts/ConfirmedStudents.py DataSources/input_updated.xlsx -HS --update=OutputFiles/master_sheet_2020_12_03_16_05_00.xlsx ```
## Changelog(03.12.2020)
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

## Pending(Script2)
