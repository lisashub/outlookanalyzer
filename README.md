# Outlook Analyzer ReadMe
[![Python 3.9](https://img.shields.io/badge/python-3.9-blue.svg)](https://www.python.org/downloads/release/python-390/)[![Python 3.10](https://img.shields.io/badge/python-3.10-blue.svg)](https://www.python.org/downloads/release/python-310/)
## Installing Dependencies
Install required packages:
> **_ANACONDA USERS:_** Sadly, Anaconda's environment defaults to an older verison of pywin32 and breaks a key .DLL path when attempting to update to versions >228. BEFORE following below pip install instructions, remove the "pywin32==[version]" line from your requirements text and save, then follow the below instructions. You will need to install the pywin32 package by entering "conda install pywin32" into your command prompt after the other packages have been installed.  See [this link](https://github.com/mhammond/pywin32/issues/1865) and [this link](https://stackoverflow.com/questions/60750197/pywin32-importerror-dll-load-failed-the-specified-module-could-not-be-found) for further details. 
> 
> If you already have pywin32 installed before working with this project, you should first uninstall Anaconda's older version using "pip uninstall pywin32" then execute "conda install pywin32".
>

(1) Open a command prompt session

(2) Navigate to the directory of the Outlook Analyzer project

(3) Enter "pip install -r requirements.txt" into your command prompt and execute

## Running script without command line options
```
python3 outlook_analyzer.py
```
Script will ask for input:
```
Max number of email messages you would like to extract (between 50 and 100000)? (Hit Enter for default: 500)
From how far back would you like to collect and analyze emails in months or days (e.g. 10m, 12d)? (Hit enter for default: 12 months ago)
What's the cutoff for the most recent emails you'd like to collect and analyze in months or days (e.g. 1m, 10d)? (Hit enter for default: today)
```

## Running script with optional command line options

You can get help with -h or --help

```
python3 outlook_analyzer.py -h
python3 outlook_analyzer.py --help
```

Format for command line options is as follows:
```
python3 outlook_analyzer.py -n <max number of email (e.g. 1000)> -s <starting point as months or days ago (e.g. 6m, 12d)> -e <end cutoff (e.g. 1m, 0d)> -O <open file after script completes - True|False> -o <out put file location>
```

Some example so you can try:
```
python3 outlook_analyzer.py -n 100
python3 outlook_analyzer.py -n 100 -s 6m
python3 outlook_analyzer.py -n 100 -s 6m -e 6d
python3 outlook_analyzer.py -n 100 -s 6m -e 6d -O True
python3 outlook_analyzer.py -n 700 -s 6m -e 6d -O False -o 'C:/Users/me/Downloads/report.pdf'
python3 outlook_analyzer.py --number 100 --start 6m --end 6d
python3 outlook_analyzer.py --start 6m --end 6d --number 100

```

If you only enter some but not all of the optional command line arguments then script will ask for the remaining arguments for number, start and end. If no command line arguments for "open" or "output" are provided then the defaults are used. Report will be saved to 'C:\WINDOWS\Temp\outlook_analyzer_report.pdf' and the pdf file will be opened after script runs.

## Understanding Script Variable Naming Standards
- Camel case is used when required for MS object model variables
- Uppercase is used for global variables (e.g. OUTPUT_NAME_STR)
- Lower case is used with an underscore between words for variables used in the script and within functions (variable_name_str)
- In general, data type is appended to the end of the variable (dict, str, int, etc) but may not be applied to all variables.
- Identations is handled with four spaces
