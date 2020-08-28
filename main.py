#!/usr/bin/env python

import os
import re
import sys
import logging
import argparse
import datetime

import init_logging

import openpyxl
import openpyxl.utils
import openpyxl.styles
import openpyxl.styles.colors
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

    # delete workbook if it exists
    output_file = config.OUTPUT_WB
    if os.path.exists(output_file):
        # eventually copy fields out...
        os.remove(output_file)

    # create a new one
    output_wb = openpyxl.Workbook()

    # handle the merged sheet
    merged_ws = output_wb.active
    merged_ws.title = config.MERGED_SHEET_NAME

    make_merged(merged_ws, vehicles_map, staff_map)

    # handle reconciliation; ok if open_rentals_files isn't there; just don't make the sheet if it isn't
    rentals_file = config.OPEN_RENTALS
    if os.path.exists(rentals_file):
        rentals_wb = openpyxl.load_workbook(rentals_file)
        rentals_ws = rentals_wb[config.OPEN_RENTALS_SHEET_NAME]

        reconciled_ws = output_wb.create_sheet(title=config.RECONCILED_SHEET_NAME)

        make_reconciled(reconciled_ws, rentals_ws, config.OPEN_RENTALS_TITLE_ROW, vehicles_ws, 1, config.OPEN_RENTALS_DRS)

    output_wb.save(output_file)



def make_merged(out_ws, vehicles_map, staff_map):
    """ generate the merged sheet from vehicles and staff """

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


def make_reconciled(reconciled_ws, rentals_ws, rentals_starting_row, vehicles_ws, vehicles_starting_row, dr_list):
    """ generate reconciled ws from rentals_ws, marking which key numbers and reservation numbers are in vehicles_ws """
    log.debug("make_reconciled: called")
    
    rentals_name_map, rentals_cols = process_title_row(rentals_ws, rentals_starting_row)
    vehicles_name_map, vehicles_cols = process_title_row(vehicles_ws, vehicles_starting_row)

    #log.debug(f"rentals_name_map: { rentals_name_map }")

    # go through the vehicles sheet and gather up all the reservation and key numbers.
    key_map = gather_column(vehicles_ws, vehicles_name_map['Key'], vehicles_starting_row, "Vehicles", 'Key')
    reservation_map = gather_column(vehicles_ws, vehicles_name_map['Reservation No'], vehicles_starting_row, "Vehicles", 'Reservation No')
    plate_map = gather_column(vehicles_ws, vehicles_name_map['Plate'], vehicles_starting_row, "Vehicles", 'Plate')

    # cell colors we'll be using
    fill_red = openpyxl.styles.PatternFill(fgColor="FFC0C0", fill_type = "solid")
    fill_green = openpyxl.styles.PatternFill(fgColor="C0FFC0", fill_type = "solid")
    fill_yellow = openpyxl.styles.PatternFill(fgColor="FFFFC0", fill_type = "solid")

    plate_column = rentals_name_map['License Plate Number']
    key_column = rentals_name_map['MVA No']
    res_column = rentals_name_map['Reservation No']
    drs_column = rentals_name_map['Cost Control No']

    log.debug("before output generation")

    # turn the list of drs into a dict for easy matching
    dr_map = {}
    for item in dr_list:
        dr_map[item] = 1

    column_dims = reconciled_ws.column_dimensions

    column_widths = {
            'Rental Region Desc': 15,
            'Rental Zone Desc': 15,
            'Rental Distict Desc': 10,
            'MVA No': 10,
            'License Plate State Code': 5,
            'License Plate Number': 10,
            'Make': 6,
            'Model': 6,
            'Ext Color Code': 6,
            'Reservation No': 15,
            'Rental Agreement No': 15,
            'CO Date': 12,
            'CO Time': 10,
            'Rental Loc Mnemonic': 15,
            'Address Line 1': 10,
            'Address Line 3': 10,
            'Full Name': 20,
            'Return Loc Mnomonic': 5,
            'Exp CI Loc Id': 5,
            'Exp CI Date': 12,
            'Exp CI Time': 10,
            'AWD Orgn Buildup Desc': 5,
            'Cost Control No': 10,
            'Booking Source Emp no': 10
            }

    for column_name in column_widths:
        if column_name in rentals_name_map:
            column_width = column_widths[column_name]
            column_index = rentals_name_map[column_name]
            # -1 in column index is because we're shifting all columns to the left since input has no column A
            column_dims[openpyxl.utils.get_column_letter(column_index-1)].width = column_width

    # specify date columns
    date_columns = [ 'CO Date', 'Exp CI Date' ]
    date_column_map = {}
    for name in date_columns:
        date_column_map[name] = 1


    # now mark the reconciled_ws with the right colors for items found/not found in vehicles
    output_row = 0
    for row in rentals_ws.iter_rows(min_row=rentals_starting_row, values_only=False):

        dr = row[drs_column -1].value

        # filter out unmatched DRs from list
        if output_row != 0 and dr not in dr_map:
            continue

        #log.debug(f"output generation row { output_row } dr { dr }")

        output_row += 1
        for cell in row:
            # skip blank first column
            if cell.column == 1:
                continue

            cell_title = rentals_cols[cell.column]

            #log.debug(f"setting row { output_row } column { cell.column } to '{ cell.value }'")

            # adjust output column by 1 because input sheet has no column 'A'
            input_column = cell.column
            output_column = input_column -1
            value = cell.value
            out_cell = reconciled_ws.cell(row=output_row, column=output_column, value=value)

            if cell_title in date_column_map:
                out_cell.number_format = 'yyyy-mm-dd'

            # now fix colors, but not for title row
            if output_row != 1:
                if input_column == res_column:
                    match_map = reservation_map
                elif input_column == key_column:
                    match_map = key_map
                    # trim leading zero from key when matching
                    value = re.sub('^0', '', value)
                elif input_column == plate_column:
                    match_map = plate_map
                else:
                    match_map = None

                if match_map != None:
                    if value in match_map:
                        out_cell.fill = fill_green
                    else:
                        out_cell.fill = fill_red

        # DEBUG ONLY
        #if output_row > 10:
        #    break




def gather_column(ws, column_num, title_row_num, table_name, column_name):
    """ gather all the values in a column into dict for easy matching """

    results = {}
    row_num = title_row_num
    for row in ws.iter_rows(min_row=title_row_num+1, min_col=column_num, max_col=column_num, values_only=True):
        row_num += 1
        cell = row[0]
        if cell == '':
            continue

        # delete formatting characters
        if isinstance(cell, str):
            cell = re.sub('[ \t-]', '', cell)

        value = str(cell)
        if value in results:

            # horrible hack to deal with n/a column in reservation no: just ignore values of n/a
            if value != 'N/A':
                # this is a duplicate entry
                log.error(f"Found duplicate for table { table_name } column { column_name }: old row { results[value] }, new row { row_num }")

        results[value] = row_num

    return results

def process_title_row(sheet, title_row_num):
    """ build two maps: one mapping column name to column number, and one from column number to name """
    title_name_map = {}
    title_cols = {}

    for row in sheet.iter_rows(min_row=title_row_num, max_row=title_row_num, values_only=False):
        for cell in row:
            #log.debug(f"title { cell.column }, { cell.row } = '{ cell.value }'")
            title_name_map[cell.value] = cell.column
            title_cols[cell.column] = cell.value
        break

    return title_name_map, title_cols

def build_map(sheet, starting_row, key_name, row_filter):
    # grab the title row
    title_name_map, title_cols = process_title_row(sheet, starting_row)

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

