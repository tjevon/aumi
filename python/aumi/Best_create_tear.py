from __future__ import print_function
import xlwings as xw
import csv
import pandas as pd
import math

import os
import logging

sch_d_df   = None
sch_ba_df  = None
sg_d_df    = None
fin_df     = None

sch_d   = "SchD_1A_1"
sch_ba  = "SchBA"
fin     = "Financials"
sg_d    = "Single_Data"
ts      = "Tear_Sheet"

years = []
companies = []

needed_sheets = [sch_d, sch_ba, fin, sg_d]

US_Fed_Tot =        ['SF04565', 'SF04566', 'SF04567', 'SF04568', 'SF04569', 'SF04570']
Oth_Gov_Tot =       ['SF10297', 'SF10298', 'SF10299', 'SF10300', 'SF10301', 'SF10302']
US_St_Ter_Tot =     ['SF10304', 'SF10305', 'SF10306', 'SF10307', 'SF10308', 'SF10309']
US_Pol_Sub_Tot =    ['SF10311', 'SF10312', 'SF10313', 'SF10314', 'SF10315', 'SF10316']
US_Spec_Rev_Tot =   ['SF10318', 'SF10319', 'SF10320', 'SF10321', 'SF10322', 'SF10323']
Corp_Tot =          ['SF10325', 'SF10326', 'SF10327', 'SF10328', 'SF10329', 'SF10330']
Hybrid_Tot =        ['SF10332', 'SF10333', 'SF10334', 'SF10335', 'SF10336', 'SF10337']
Par_Sub_Aff_Tot =   ['SF02189', 'SF02190', 'SF02191', 'SF02192', 'SF02193', 'SF02194']
Gov_Ag_Tot =        [US_Fed_Tot, Oth_Gov_Tot]
Muni_Tot =          [US_St_Ter_Tot, US_Pol_Sub_Tot, US_Spec_Rev_Tot]

Combo_Asset_Classes = [Gov_Ag_Tot, Muni_Tot]

Asset_Class_Tots = [US_Fed_Tot, Oth_Gov_Tot, US_St_Ter_Tot, US_Pol_Sub_Tot, Gov_Ag_Tot, Corp_Tot, US_Spec_Rev_Tot, Hybrid_Tot, Par_Sub_Aff_Tot, Muni_Tot]
LINES_PER_ASSET_CLASS = 10

data_dir = os.getenv('DATAPATH', './')
log_dir = os.getenv('LOGPATH', './')

def init_logger():
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

def read_sch_d(entity_id, csv_filename):
    global sch_d_df
    sch_d_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2)
    for i in sch_d_df.iloc[1]:
        if pd.notnull(i):
            years.append(int(i))
    years.sort()
    # should read df with group # for column lable and field # for row label
    sch_d_df = pd.read_csv(csv_filename, header=3, index_col=0)
    for i in sch_d_df.columns:
        if i.find("Unnamed") != -1:
            continue
        if i.find(".") != -1:
            continue
        if pd.notnull(i):
            companies.append(i)
    pass

def read_fin(entity_id, csv_filename):
    global fin_df
    sch_d_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2)
    for i in sch_d_df.iloc[1]:
        if pd.notnull(i):
            years.append(int(i))
    years.sort()
    # should read df with group # for column lable and field # for row label
    sch_d_df = pd.read_csv(csv_filename, header=3, index_col=0)
    for i in sch_d_df.columns:
        if i.find("Unnamed") != -1:
            continue
        if i.find(".") != -1:
            continue
        if pd.notnull(i):
            companies.append(i)
    pass

def read_csvs( xl_wb, entity_id, file_names ):
    logger.debug("Enter")
    global sch_ba_df
    global sch_d_df
    global fin_df
    global sg_d_df

    for file_name in file_names:
        csv_filename = data_dir + "\\" + file_name
        if file_name.find(sch_d) != -1:
            read_sch_d(entity_id, csv_filename)
        elif file_name.find(fin) != -1:
            read_fin(entity_id, csv_filename)
    pass
    #fin_df = pd.read_csv(csv_filename)
#        elif file_name.find(sch_ba) != -1:
#            sch_ba_df = pd.read_csv(csv_filename)
#        elif file_name.find(sg_d) != -1:
#            sg_d_df = pd.read_csv(csv_filename)

def build_4_sheets(xl_wb):
    xl_sheet = xl_wb.sheets(sch_d)
    xl_range = xl_sheet.range('D6:P108')
    xl_range.clear_contents()
    for co in companies:
        build_sch_d(xl_sheet, co)

def build_sch_d(xl_sheet, co):

    ig_tot = 0
    hy_tot = 0
    num_years = len(years)
    num_cols = (num_years - 1) + num_years + 2
    num_rows = len(Asset_Class_Tots) * LINES_PER_ASSET_CLASS + 3
    column_heading = []
    # allocate 2 D list for numeric data
    sch_d_mat = [[0.0 for x in range(num_cols)] for y in range(num_rows)]

    col_idx = 0
    year_idx = 0
    count = 0
    while col_idx < num_cols and year_idx < num_years:
        if col_idx < 2 or (col_idx % 2) == 1 :
            column_heading.append(years[year_idx])
            year_idx += 1
        else:
            tmp_str = "%s%% Chg" % (years[year_idx-1])
            column_heading.append(tmp_str)
        col_idx += 1

    tmp_str = "%s%% Chg" % (years[year_idx-1])
    column_heading.append(tmp_str)
    tmp_str = "%d Yr %% Chg" % (num_years-1)
    column_heading.append(tmp_str)
    tmp_str = "Annualized % Chg"
    column_heading.append(tmp_str)

    # for each year (col), identify the row label and column label for the DataFrame
    mat_col_idx = 0
    year_idx = 0
    while year_idx < len(years) and mat_col_idx < num_cols:
        col_label = None
        if year_idx == 0:
            col_label = co
        else:
            col_label = "%s.%d" % (co, year_idx)
        mat_line_idx = 0

        if mat_col_idx > 0 and (mat_col_idx % 2) == 0:
            # leave blank columns to be filled in with calculated values
            mat_col_idx += 1
            continue

        # for each subsection
        for asset_class in Asset_Class_Tots:
            asset_class_ig_tot = 0
            asset_class_hy_tot = 0
            # if the subsection is a combination of others subsections
            if asset_class in Combo_Asset_Classes:
                tmp_mat_line_idx = mat_line_idx
                for component_asset_class in asset_class:
                    mat_line_idx = tmp_mat_line_idx

                    for asset_class_label, count in zip(component_asset_class, range(0,len(component_asset_class))):
                        sch_d_mat[mat_line_idx][mat_col_idx] += float(sch_d_df.loc[asset_class_label, col_label])
                        if count < 3:
                            asset_class_ig_tot += int(sch_d_mat[mat_line_idx][mat_col_idx])
                        elif count < 6:
                            asset_class_hy_tot += int(sch_d_mat[mat_line_idx][mat_col_idx])
                        mat_line_idx += 1
            else:
                # for each line in each subsecion
                for asset_class_label, count in zip(asset_class, range(0,len(asset_class))):
                    sch_d_mat[mat_line_idx][mat_col_idx] = sch_d_df.loc[asset_class_label, col_label]
                    if count < 3:
                        asset_class_ig_tot += int(sch_d_mat[mat_line_idx][mat_col_idx])
                    elif count < 6:
                        asset_class_hy_tot += int(sch_d_mat[mat_line_idx][mat_col_idx])
                    mat_line_idx += 1

            # these are incorrect - double counting on some asset classes
            sch_d_mat[mat_line_idx][mat_col_idx] = asset_class_ig_tot
            mat_line_idx += 1
            sch_d_mat[mat_line_idx][mat_col_idx] = asset_class_hy_tot
            mat_line_idx += 1
            sch_d_mat[mat_line_idx][mat_col_idx] = asset_class_ig_tot + asset_class_hy_tot
            mat_line_idx += 2
            ig_tot += asset_class_ig_tot
            hy_tot += asset_class_hy_tot

        sch_d_mat[mat_line_idx][mat_col_idx] = ig_tot
        mat_line_idx += 1
        sch_d_mat[mat_line_idx][mat_col_idx] = hy_tot
        mat_line_idx += 1
        sch_d_mat[mat_line_idx][mat_col_idx] = ig_tot + hy_tot
        mat_line_idx += 1

        mat_col_idx += 1
        year_idx += 1
        pass

    # fill in the blank columns with % change data
    # fill in the year over year change
    mat_col_idx = 2
    prev_year_idx = 0
    cur_year_idx = 1
    while mat_col_idx < (num_cols - 2):
        for mat_line_idx in range(0,num_rows):
            if float(sch_d_mat[mat_line_idx][prev_year_idx]) > 0.0:
                sch_d_mat[mat_line_idx][mat_col_idx] = float(sch_d_mat[mat_line_idx][cur_year_idx])/float(sch_d_mat[mat_line_idx][prev_year_idx]) - 1.0
        prev_year_idx = cur_year_idx
        cur_year_idx += 2
        mat_col_idx += 2

    # fill in the 5 year change and annualized change
    cur_year_idx = num_cols - 4
    for mat_line_idx in range(0,num_rows):
        mat_col_idx = num_cols - 2
        if float(sch_d_mat[mat_line_idx][1]) > 0.0:
            sch_d_mat[mat_line_idx][mat_col_idx] = float(sch_d_mat[mat_line_idx][cur_year_idx])/float(sch_d_mat[mat_line_idx][1]) - 1.0
        mat_col_idx += 1
        sch_d_mat[mat_line_idx][mat_col_idx] = math.pow(1.0 + float(sch_d_mat[mat_line_idx][mat_col_idx-1]),(1.0/5.0)) - 1.0
    pass

    # push it out to the spreadsheet
    attempt = 0
    while True:
        try:
            xl_range = xl_sheet.range('B1')
            xl_range.value = co

            xl_range = xl_sheet.range('D4')
            xl_range.value = column_heading

            xl_range = xl_sheet.range('D6')
            xl_range.value = sch_d_mat
            break
        except:
            logger.error("Unable to write to excel workbook: %d", attempt)
            attempt += 1
            if attempt > 5:
                return

def read_sg_d(xl_wb, file_name):
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

    xl_sheet = xl_wb.sheets(sheet_name)
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

#    xl_wb = xw.Book.caller()
    xl_wb = xw.Book(r'C:\Users\tjevon\Dropbox\code\sandbox\iaum\xl\TearSheet_py2.xlsm')

    file_info = get_file_dict()
    for entity_id in file_info:
        logger.info("Industry: %s", entity_id)
        for file in file_info[entity_id]:
            logger.info("File: %s", file)

    # findout what sheets are already open in the workbook
    available_sheets = []
    for i in xl_wb.sheets:
        logger.debug(i.name)
        available_sheets.append(i.name)
    logger.debug("Have Sheets:")

    for sheet_name in needed_sheets:
        if sheet_name not in available_sheets:
            xl_wb.sheets.add(sheet_name)
        xl_sheet = xl_wb.sheets(sheet_name)
#        if sheet_name != ts:
#            xl_sheet.clear_contents()

    # for each company, read in the csv's and calculate the tear sheet
    for entity_id in file_info:
        read_csvs(xl_wb, entity_id, file_info[entity_id])
        build_4_sheets(xl_wb)
#        xl_mac = xl_wb.macro('CalcSheet')
#        xl_mac(ts)
    logger.info("Leave")

################################################################################
if __name__ == '__main__':
    create_tearsheets()
