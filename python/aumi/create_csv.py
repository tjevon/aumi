from __future__ import print_function
from xlwings import Book, Range
import struct
import pandas as pd

import os
import logging
from sys import exc_info, stderr, exit
from optparse import OptionParser

def create_csv():
    log_dir = os.getenv('LOGPATH', './')
    logger = logging.getLogger('onerun')
    logger.setLevel(logging.DEBUG)
    data_dir = os.getenv('DATAPATH', './')

    fullpath = ("%s/create_csv.log" % (log_dir))
    fh = logging.FileHandler(fullpath)
    fh.setLevel(logging.INFO)

    log_format = logging.Formatter('[%(asctime)s %(levelname)s %(module)s::%(funcName)s %(lineno)d] %(message)s')
    fh.setFormatter(log_format)

    logger.addHandler(fh)
    logger.info("Enter")

    xl_wb = Book.caller()
    create_csv_from(logger, data_dir, xl_wb, "SchD_1A_1", True)
    create_csv_from(logger, data_dir, xl_wb, "SchBA", True)
    create_csv_from(logger, data_dir, xl_wb, "Financials", True)
    create_csv_from(logger, data_dir, xl_wb, "Single_Data", True)
    create_csv_from(logger, data_dir, xl_wb, "Prospect_List", False)
    logger.info("Leave")

def create_csv_from(logger, data_dir, xl_wb, sheet_name, include_entity):
    logger.debug("Enter")
    logger.info("Process - %s", sheet_name)

    xl_sheet = xl_wb.sheets[sheet_name]
    entity_id = xl_sheet.range('B1').value
    empty_line = 0
    bottom = -1
    right = 0
    x = 1
    while x < 10000:
        cell = "A%d" % x
        logger.debug( "%d %s %d %d %d", x, cell, bottom, right, empty_line)
        test_range = xl_sheet.range(cell)
        test_range = test_range.current_region
        last_cell = test_range.last_cell
        if test_range.count > 1 or last_cell.value:
            # region bigger than 1 or single cell that contains data
            if x - 1 == bottom:
                # must be an empty line following a valid region, ignore it
                x += 1
                continue

            if right < last_cell.column:
                right = last_cell.column
            bottom = test_range.last_cell.row
            x = bottom + 1
            empty_line = 0
        else:
            empty_line += 1
            x += 1
        if empty_line > 10:
            break

    top_range = xl_sheet.range('A1')
    top_range = top_range.resize(bottom, right)
    df = pd.DataFrame(top_range.value)
    if include_entity:
        sheet_filename = data_dir + "\\" + entity_id + "_" + sheet_name + ".csv"
    else:
        sheet_filename = data_dir + "\\" + sheet_name + ".csv"
    logger.debug("filename %s", sheet_filename)
    df.to_csv(sheet_filename,index=False, header=False)

    logger.debug("Leave")
