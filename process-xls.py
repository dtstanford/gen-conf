#!/usr/bin/env python3
 
import argparse
import os
import sys
import openpyxl

# Set some default global vars for use in argparse
WS_NAME = 'ACL REQUEST FORM'
START_CELL='A3'

# Set some global vars
CWD = os.getcwd()

def parse_args():
    parser = argparse.ArgumentParser(description='A command-line tool for Static1 processing of a \
        particular client\'s ACL request forms.',)

    parser.add_argument('file', type=str, help='The .xlsx file to be processed.')
    parser.add_argument('acl_name', type=str, help='The name of the firewall ACL this will be \
                        applied to.')
    parser.add_argument('--sheet', type=str, help='The worksheet containing the ACL request form. \
                        (default: \'ACL REQUEST FORM\')', default=WS_NAME)
    parser.add_argument('--outfile', type=str, help='Write output to a file in addition to the \
                        screen. (default: \'outfile.txt\')', default='outfile.txt')
    parser.add_argument('--start-cell', type=str, help='The upper-leftmost worksheet cell to \
                        begin processing ACL rules from (e.g., \'B4\'). (default: A3)',
                        default=START_CELL)

    args = parser.parse_args(['eqfwchange.xlsx', 'aclname'])

    return args

def validate_start_cell(start_cell):
    def error_and_exit():
        print()
        print('The supplied start-cell argument (\'{start_cell}\') is not supported.'
            .format(start_cell=start_cell))
        print()
        print('    ACCEPTABLE START-CELL RANGE: A1 -> A99')
        print()

        sys.exit(1)

    if start_cell[0].isalpha() and start_cell[1:].isnumeric():
        format_correct = True
    else:
        format_correct = False

    if 2 <= len(start_cell) <= 3:
        length_correct = True
    else:
        length_correct = False

    if format_correct and length_correct:
        return start_cell
    else:
        error_and_exit()

def try_load_workbook(file):
    try:
        wb = openpyxl.load_workbook(filename=file)
    except openpyxl.utils.exceptions.InvalidFileException:
        print()
        print('The input file you supplied (\'{file}\') doesn\'t have a valid extension:'
            .format(file=file))
        print()
        print('    VALID EXTENSIONS: .xlsx, .xlsm, .xltx, .xltm')
        print()

        sys.exit(1)
    except FileNotFoundError:
        print()
        print('Unable to locate the supplied file in the current directory.')
        print()
        print('    Attempted to open: {cwd}/{file}'.format(cwd=CWD, file=file))
        print()

        sys.exit(1)

    return wb

def try_load_worksheet(wb, ws_name, wb_name):
    try:
        ws = wb[ws_name]
    except KeyError:
        print()
        print('The worksheet \'{worksheet}\' doesn\'t exist in the input file \'{file}\'.'\
            .format(worksheet=ws_name, file=wb_name))
        print()

        sys.exit(1)

    return ws

def convert_alpha_to_num(letter):
    return ord(letter.lower()) - 96


def main():
    args = parse_args()
    file = args.file
    acl = args.acl_name
    ws_name = args.sheet
    start_cell = validate_start_cell(args.start_cell)
    start_col = convert_alpha_to_num(start_cell[0])
    start_row = int(start_cell[1:])

    wb = try_load_workbook(file)
    ws = try_load_worksheet(wb, ws_name, file)

    src_net_col = start_col
    src_net_desc_col = start_col + 1
    dest_net_col = start_col + 2
    ip_proto_col = start_col + 4
    ip_proto_port_col = start_col + 5

    for row in ws.iter_rows(min_col=start_col, min_row=start_row, values_only=True):
        print(row[src_net_col - 1])
        break

if __name__ == '__main__':
    main()