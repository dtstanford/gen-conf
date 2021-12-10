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
*Windows*: https://www.python.org/downloads/release/python-3100

*Mac OS*:

Python 3 comes pre-installed on Mac OS. Confirm you have it in your system PATH:

```
which python3
```

### Install Git
*Windows*: https://git-scm.com/download/win

*Mac OS*:  
```
xcode-select --install
```

### Download Script Repository  
For both Windows and Mac OS:  
```
git clone https://github.com/dtstanford/gen-conf.git
```

### Install Required Python Package(s)  
From within the script directory:  
```
python3 -m pip install -r requirements.txt
```

## Running the Script
1. Ensure you have your .xlsx file in the same directory as the script.
2. Run the script, giving it the .xlsx file name and the name of the ACL you want to modify:

Example: `./gen-conf.py changes.xlsx outside_in_acl`

### Additional Options
For a list of full options, run the script using the `-h` (help) argument:

```
✗ ./gen-conf.py -h
usage: gen-conf.py [-h] [--sheet SHEET] [--vl-sheet VL_SHEET] [--outfile OUTFILE] [--start-cell START_CELL] file acl_name

A command-line tool for Static1 processing of a particular client's ACL request forms.

positional arguments:
  file                  The .xlsx file to be processed.
  acl_name              The name of the firewall ACL this will be applied to.

options:
  -h, --help            show this help message and exit
  --sheet SHEET         The worksheet containing the ACL request form. (default: 'ACL REQUEST FORM')
  --vl-sheet VL_SHEET   The worksheet containing the VLOOKUP referenced data. (default: 'Data')
  --outfile OUTFILE     Write output to a file in addition to the screen.
  --start-cell START_CELL
                        The upper-leftmost worksheet cell to begin processing ACL rules from (e.g., 'B4'). (default: A3)
```

### Example Output  
```
# ./gen-conf.py fw-changes-example.xlsx outside_acl --date 01011970

                               ***NOTICE***
This script is intended to simplify the ACL change request process. However,
it is the user's responsibility to validate the configuration prior to final
implementation. USE WITH CAUTION.

Continue? [Y/n]: Y
------------------------------

object-group network CompanyX-Clients_01-01-1970
  network-object host 192.168.195.53
  network-object 192.168.97.0 255.255.255.192

object-group network Big-Org-Initiators_01-01-1970
  network-object host 10.12.158.186

access-list outside_acl line 1 extended permit tcp object-group CompanyX-Clients_01-01-1970 host 172.20.240.132 eq 22443

access-list outside_acl line 1 extended permit tcp object-group Big-Org-Initiators_01-01-1970 host 172.18.240.131 eq 22
access-list outside_acl line 1 extended permit tcp object-group Big-Org-Initiators_01-01-1970 host 172.18.240.131 range 500 599

---- BACKOUT CONFIG ----
no access-list outside_acl extended permit tcp object-group CompanyX-Clients_01-01-1970 host 172.20.240.132 eq 22443
no access-list outside_acl extended permit tcp object-group Big-Org-Initiators_01-01-1970 host 172.18.240.131 eq 22
no access-list outside_acl extended permit tcp object-group Big-Org-Initiators_01-01-1970 host 172.18.240.131 range 500 599
no object-group network CompanyX-Clients_01-01-1970
no object-group network Big-Org-Initiators_01-01-1970
```