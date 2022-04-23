# Outlook Analyzer ReadMe

## Installing Dependencies
Install required packages:

(1) Open command prompt

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
python3 outlook_analyzer.py -n <max number of email (e.g. 1000)> -s <starting point as months or days ago (e.g. 6m, 12d)> -e <end cutoff (e.g. 1m, 0d)>
```

Some example so you can try:
```
python3 outlook_analyzer.py -n 100
python3 outlook_analyzer.py -n 100 -s 6m
python3 outlook_analyzer.py -n 100 -s 6m -e 6d
python3 outlook_analyzer.py --number 100 --start 6m --end 6d
python3 outlook_analyzer.py --start 6m --end 6d --number 100
```

If you only enter some but not all of the optional command line arguments then script will ask for the remaining arguments.

## Understanding Script Variable Naming Standards
- Camel case is used when required for MS object model variables
- Uppercase is used for global variables (e.g. OUTPUT_NAME_STR)
- Lower case is used with an underscore between words for variables used in the script and within functions (variable_name_str)
- Data type is appended to the end of the variable (dict, str, int, etc).
- Identations is handled with four spaces
