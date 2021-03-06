from __future__ import print_function
import xlwings as xw
import csv

import os
import logging

sch_d = "SchD_1A_1"
sch_ba = "SchBA"
fin = "Financials"
sg_d = "Single_Data"
ts = "Tear_Sheet"
needed_sheets = [sch_d, sch_ba, fin, sg_d]

data_dir = os.getenv('DATAPATH', './')

def init_logger():
    log_dir = os.getenv('LOGPATH', './')
    logger = logging.getLogger('onerun')
    logger.setLevel(logging.DEBUG)

    fullpath = ("%s/create_tearsheets.log" % (log_dir))
    fh = logging.FileHandler(fullpath)
    fh.setLevel(logging.DEBUG)

    log_format = logging.Formatter('[%(asctime)s %(levelname)s %(module)s::%(funcName)s %(lineno)d] %(message)s')
    fh.setFormatter(log_format)

    logger.addHandler(fh)
    logger.info("Leave")
    return logger

logger = init_logger()

def get_filenames():
    logger.info("Enter")
    logger.info("DATAPATH: %s", data_dir)

    files = os.listdir(data_dir)
    file_names = []
    for file in files:
        logger.info("file: %s", file)
        if file.find(sch_d) != -1:
            file_names.append(file)
        elif file.find(sch_ba) != -1:
            file_names.append(file)
        elif file.find(fin) != -1:
            file_names.append(file)
        elif file.find(sg_d) != -1:
            file_names.append(file)
    logger.info("Leave")
    return file_names

def get_file_dict():
    file_names = get_filenames()
    file_dict = {}

    for file in file_names:
        idx = file.find('_')
        entity_id = file[:idx]
        if entity_id in file_dict:
            file_dict[entity_id].append(file)
        else:
            file_dict[entity_id] = [file]
    return file_dict

def read_csvs( xl_wb, entity_id, file_names ):
    logger.debug("Enter")

    for file_name in file_names:
        csv_filename = data_dir + "\\" + file_name

        # get rid of company id
        idx = file_name.find('_')
        sheet_name = file_name[idx+1:]
        # get rid of .csv
        idx = sheet_name.find(".")
        sheet_name = sheet_name[:idx]

        logger.info("Process csv_filename: %s sheet_name: %s", csv_filename, sheet_name)
        try:
            csvfile = open( csv_filename, 'rb')
        except:
            logger.error("Caught open for file %s", csv_filename)
            return

        try:
            reader = csv.reader(csvfile, delimiter=',', dialect='excel')
        except:
            logger.error("Caught csv.reader for file %s", csv_filename)
            return

        xl_sheet = xw.sheets(sheet_name)
        xl_sheet.clear_contents()

        my_matrix = []
        for row in reader:
            my_matrix.append(row)
        xl_range = xl_sheet.range('A1')
        xl_range.value = my_matrix
    logger.debug("LEAVE:")

#Entry point
def create_tearsheets():
    logger.info("Enter")
    xl_wb = xw.Book.caller()

    file_info = get_file_dict()
    for entity_id in file_info:
        logger.info("Company: %s", entity_id)
        for file in file_info[entity_id]:
            logger.info("File: %s", file)

    # findout what sheets are already open in the workbook
    available_sheets = []
    for i in xw.sheets:
        logger.debug(i.name)
        available_sheets.append(i.name)
    logger.debug("Have Sheets:")

    for sheet_name in needed_sheets:
        if sheet_name not in available_sheets:
            xw.sheets.add(sheet_name)
        xl_sheet = xw.sheets(sheet_name)
        if sheet_name != ts:
            xl_sheet.clear_contents()

    # for each company, read in the csv's and calculate the tear sheet
    for entity_id in file_info:
        read_csvs(xl_wb, entity_id, file_info[entity_id])
        xl_mac = xl_wb.macro('CalcSheet')
        xl_mac(ts)
    logger.info("Leave")



