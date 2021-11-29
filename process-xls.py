#!/usr/bin/env python3
 
import argparse
from datetime import date
import ipaddress
import os
import re
import sys
import openpyxl

# Set some default global vars for use in argparse
WS_NAME = 'ACL REQUEST FORM'
START_CELL='A3'

# Set some global vars
CWD = os.getcwd()

def print_notice():
    notice =  '                                   *NOTICE*                                   \n'
    notice += 'This script is intended to simplify the ACL change request process. However, \n'
    notice += 'it is the user\'s responsibility to validate the configuration prior to final \n'
    notice += 'implementation. USE WITH CAUTION.'

    print()
    print(notice)
    print()
    response = input('Continue? [Y/n]: ')
    print()

    if response == '' or response.lower() == 'y':
        return
    else:
        sys.exit(0)

def parse_args():
    parser = argparse.ArgumentParser(description='A command-line tool for Static1 processing of a \
        particular client\'s ACL request forms.',)

    parser.add_argument('file', type=str, help='The .xlsx file to be processed.')
    parser.add_argument('acl_name', type=str, help='The name of the firewall ACL this will be \
                        applied to.')
    parser.add_argument('--sheet', type=str, help='The worksheet containing the ACL request form. \
                        (default: \'ACL REQUEST FORM\')', default=WS_NAME)
    parser.add_argument('--outfile', type=str, help='Write output to a file in addition to the \
                        screen.')
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

def parse_to_list(string):
    split_string = re.split(r',|\n', string)
    split_string[:] = [elem.strip() for elem in split_string]

    return split_string

def generate_net_desc(net_desc):
    net_desc = net_desc[0:53]
    net_desc = net_desc.replace(' ', '-')
    todays_date = date.today().strftime('%m-%d-%Y')
    net_desc = net_desc + '_' + todays_date

    return net_desc


def try_generate_net(net):
    try:
        net =ipaddress.IPv4Network(net)
    except:
        print()
        print('There was an error attempting to process the host/network: {net}'.format(net=net))
        print()

        sys.exit(1)
    
    return net

def generate_objgrp_config(nets, nets_desc):
    host_list = []
    net_list = []
    
    for net in nets:
        net = try_generate_net(net)
        if net.num_addresses == 1:
            host = str(net.hosts()[0])
            host_list.append(host)
        else:
            net = net.with_netmask.replace('/', ' ')
            net_list.append(net)
    
    nets_desc = generate_net_desc(nets_desc)

    objgrp_config = 'object-group network ' + nets_desc + '\n'
    
    for host in host_list:
        objgrp_config += '  network-object host ' + host + '\n'
    
    for net in net_list:
        objgrp_config += '  network-object ' + net + '\n'

    return objgrp_config

def main():
    print_notice()
    args = parse_args()
    file = args.file
    acl = args.acl_name
    ws_name = args.sheet
    outfile = args.outfile
    start_cell = validate_start_cell(args.start_cell)
    start_col = convert_alpha_to_num(start_cell[0])
    start_row = int(start_cell[1:])

    wb = try_load_workbook(file)
    ws = try_load_worksheet(wb, ws_name, file)

    src_nets_col = start_col
    src_net_desc_col = start_col + 1
    dest_nets_col = start_col + 2
    ip_proto_col = start_col + 4
    ip_proto_ports_col = start_col + 5

    for row in ws.iter_rows(min_col=start_col, min_row=start_row, values_only=True):
        src_nets_cell = row[src_nets_col - 1]
        src_net_desc_cell = row[src_net_desc_col - 1]
        dest_nets_cell = row[dest_nets_col - 1]
        ip_proto_cell = row[ip_proto_col - 1]
        ip_proto_ports_cell = row[ip_proto_ports_col - 1]
        
        src_nets_cell = parse_to_list(src_nets_cell)
        src_nets_objgrp_config = generate_objgrp_config(src_nets_cell, src_net_desc_cell)

        print(src_nets_objgrp_config)

if __name__ == '__main__':
    main()