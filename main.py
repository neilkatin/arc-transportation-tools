#!/usr/bin/env python

import os
import sys
import logging
import argparse
import datetime

import init_logging

import openpyxl
import openpyxl.utils
import dotenv

import config as config_static


log = logging.getLogger(__name__)


DATESTAMP = datetime.datetime.now().strftime("%Y-%m-%d")

class AttrDict(dict):
    def __init__(self, *args, **kwargs):
        super(AttrDict, self).__init__(*args, **kwargs)
        self.__dict__ = self

def rentals_filter(row):
    return row['Ctg'] == 'R' and row['Key'] == ''

def empty_filter(row):
    return True


def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config_dotenv = dotenv.dotenv_values(verbose=True)

    config = AttrDict()
    for item in dir(config_static):
        if not item.startswith('__'):
            config[item] = getattr(config_static, item)

    #log.debug(f"config after copy: { config.keys() }")

    for key, val in config_dotenv.items():
        config[key] = val

    vehicles_file = config.VEHICLES
    if not os.path.exists(vehicles_file):
        log.fatal(f"Vehicles file { vehicles_file } not found")
        sys.exit(1)

    staff_file = config.STAFF_ROSTER
    if not os.path.exists(vehicles_file):
        log.fatal(f"Staff Roster file { staff_file } not found")
        sys.exit(1)

    vehicles_wb = openpyxl.load_workbook(vehicles_file)
    vehicles_ws = vehicles_wb[config.VEHICLES_SHEET_NAME]
    vehicles_map = build_map(vehicles_ws, 1, "Rcvd From", rentals_filter)

    staff_wb = openpyxl.load_workbook(staff_file)
    staff_ws = staff_wb[config.STAFF_ROSTER_SHEET_NAME]
    staff_map = build_map(staff_ws, 4, "Name", empty_filter)

    merged_file = config.MERGED
    if os.path.exists(merged_file):
        # eventually copy fields out...
        os.remove(merged_file)

    merged_wb = openpyxl.Workbook()
    merged_ws = merged_wb.active
    merged_ws.title = config.MERGED_SHEET_NAME

    combine_data(merged_ws, vehicles_map, staff_map)

    merged_wb.save(merged_file)

def combine_data(out_ws, vehicles_map, staff_map):

    # first make the title row
    vehicles_columns = [ 'Name', 'Reservation No', 'GAP', 'Date Received', ]
    vehicles_col_width = [ 20,    15,              15,     12, ]
    staff_columns = [ 'Email', 'Cell phone', 'Assigned', 'Checked in', 'Supervisor(s)', ]
    staff_col_width = [ 30,    14,              20,        20,           20, ]


    column_map = {}
    column_index = 0
    col_dims = out_ws.column_dimensions
    for i, name in enumerate(vehicles_columns):
        column_index += 1
        column_map[name] = { 'index': column_index, 'map': vehicles_map, 'col_name': name, 'map_name': 'Vehicles' }
        out_ws.cell(row=1, column=column_index, value=name)
        col_dims[openpyxl.utils.get_column_letter(column_index)].width = vehicles_col_width[i]

    for i, name in enumerate(staff_columns):
        column_index += 1
        column_map[name] = column_index
        column_map[name] = { 'index': column_index, 'map': staff_map, 'col_name': name, 'map_name': 'Staff Roster' }
        out_ws.cell(row=1, column=column_index, value=name)
        col_dims[openpyxl.utils.get_column_letter(column_index)].width = staff_col_width[i]

    # the name in the vehicles_map is 'Rcvd From'
    column_map['Name']['col_name'] = 'Rcvd From'

    # now generate the data
    keys = sorted(vehicles_map.keys(), key=str.lower)

    rownum = 1
    for row_name in keys:
        row = vehicles_map[row_name]
        rownum += 1
        for name, data in column_map.items():
            out_index = data['index']
            in_map = data['map']
            in_col = data['col_name']

            if row_name not in in_map:
                continue

            entry = in_map[row_name]

            if in_col not in entry:
                log.error(f"can't find key '{ in_col }' in map '{ data['map_name'] }'")
            else:
                value = entry[in_col]
                out_ws.cell(row=rownum, column=out_index, value=value)



def build_map(sheet, starting_row, key_name, row_filter):
    # grab the title row
    for row in sheet.iter_rows(min_row=starting_row, max_row=starting_row, values_only=False):
        title_row = row
        title_name_map = {}
        title_cols = {}
        for cell in title_row:
            #log.debug(f"title { cell.column }, { cell.row } = '{ cell.value }'")
            title_name_map[cell.value] = cell.column
            title_cols[cell.column] = cell.value
        break

    #log.debug(f"title_name_map: { title_name_map }")

    results = {}
    for row in sheet.iter_rows(min_row=starting_row+1, values_only=False):
        row_map = {}
        for cell in row:
            value = cell.value
            title = title_cols[cell.column]
            row_map[title] = value

        if not row_filter(row_map):
            continue

        #log.debug(f"row_map: { row_map }")
        key = row_map[key_name]

        if key in results:
            raise Exception(f"Error: key { key } already in results dictionary")

        results[key] = row_map

    return results




def parse_args():
    parser = argparse.ArgumentParser(
            description="process support for the regional bootcamp mission card system",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")

    #group = parser.add_mutually_exclusive_group(required=True)
    #group.add_argument("-p", "--prod", "--production", help="use production settings", action="store_true")
    #group.add_argument("-d", "--dev", "--development", help="use development settings", action="store_true")

    args = parser.parse_args()
    return args


if __name__ == "__main__":
    main()

