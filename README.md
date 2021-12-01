# <span>gen-conf.py</span>

## Description
This is a fairly simple script meant to generate specific Cisco ASA ACL configuration statements 
based off an input .xlsx file. It was written with a particular client and .xlsx file format 
in mind, and therefore does have some inherent limitations. However, overall, is should still speed
up any routine processing of these particular .xlsx ACL requests.

**Proof-read every generated configuration. Use with caution.**

## Installation
This script was written on Mac OS, but should still run on any operating system with Python 3
and the required pip package(s) installed.

### Install Python
__Windows__: https://www.python.org/downloads/release/python-3100/

__Mac OS__:

Python 3 comes pre-installed on Mac OS. Confirm you have it in your system PATH:

`which python3`  
`python3 -V`

### Install Git
(Needs to be populated...)

### Download Script Repository
(Needs to be populated...)

### Install Required Python Package(s)
`python3 -m pip install -r requirements.txt`

## Running the Script
1. Ensure you have your .xlsx file in the same directory as the script.
2. Run the script, giving it the .xlsx file name and the name of the ACL you want to modify:

Example: `./gen-conf.py changes.xlsx outside_in_acl`

### Additional Options
For a list of full options, run the script using the `-h` (help) argument:

Example: `./gen-conf.py -h`