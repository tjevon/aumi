from __future__ import print_function
import xlwings as xw
import csv
import pandas as pd
import math

import os
import logging
import Best_fid_defs as best


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
companies = set()
company_map = {}

needed_sheets = [sch_d, sch_ba, fin, sg_d]

data_dir = os.getenv('DATAPATH', './')
log_dir = os.getenv('LOGPATH', './')

def init_logger():
    logger = logging.getLogger('onerun')
    logger.setLevel(logging.DEBUG)

    fullpath = ("%s/Best_create_tearsheets.log" % (log_dir))
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
    logger.info("Enter")
    file_names = get_filenames()
    file_dict = {}

    for file in file_names:
        idx = file.find('_')
        entity_id = file[:idx]
        if entity_id in file_dict:
            file_dict[entity_id].append(file)
        else:
            file_dict[entity_id] = [file]
    logger.info("Leave")
    return file_dict

def get_column_headings(years, num_cols):
    column_heading = []
    col_idx = 0
    year_idx = 0
    num_years = len(years)
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
    return column_heading

def get_col_label(year_idx, fid):
    col_label = None
    if year_idx == 0:
        col_label = fid
    else:
        col_label = "%s.%d" % (fid, year_idx)
    return col_label

def read_sg_d(entity_id, csv_filename):
    logger.info("Enter: %s", csv_filename)
    global companies
    the_df = None
    the_df = pd.read_csv(csv_filename, header=2, index_col=0, dtype='unicode')
    logger.info("Leave")
    return the_df

def read_vert(entity_id, csv_filename):
    logger.info("Enter: %s", csv_filename)
    global years
    global companies
    the_df = None
    years = []
    tmp_years = []
    the_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode')
    for i in the_df.iloc[1]:
        if pd.notnull(i):
            tmp_years.append(int(i))
    years = sorted(set(tmp_years))
    # should read df with field # for column lable and group # for row label
    the_df = pd.read_csv(csv_filename, header=3, index_col=0, dtype='unicode')
    for i in the_df.index:
        if i.find("Unnamed") != -1:
            continue
        if i.find(".") != -1:
            continue
        if i.find("AMB#") != -1:
            continue
        if pd.notnull(i):
            companies.add(i)
    for i in the_df.columns:
        if i.find('Calc') != -1:
            the_df[i].replace(regex=True,inplace=True,to_replace=r'%',value=r'')

    logger.info("Leave")
    return the_df
    csv_filename = data_dir + "\\" + file_name

def read_csv_all(xl_wb, entity_id, file_names):
    logger.info("Enter")
    global sg_d_df
    for file_name in file_names:
        csv_filename = data_dir + "\\" + file_name
        if file_name.find(sg_d) != 1:
            sg_d_df = read_sg_d(entity_id, csv_filename)

    logger.info("Leave")
    return

def read_csvs( xl_wb, entity_id, file_names ):
    logger.info("Enter")
    global sch_ba_df
    global sch_d_df
    global fin_df

    for file_name in file_names:
        csv_filename = data_dir + "\\" + file_name
        if file_name.find(sch_d) != -1:
            sch_d_df = read_vert(entity_id, csv_filename)
        elif file_name.find(fin) != -1:
            fin_df = read_vert(entity_id, csv_filename)
        elif file_name.find(sch_ba) != -1:
            sch_ba_df = read_vert(entity_id, csv_filename)
    pass
    logger.info("Leave")
    return

def calc_percent_columns(the_mat):
    num_rows = len(the_mat)
    num_cols = len(the_mat[0])
    mat_col_idx = 2
    prev_year_idx = 0
    cur_year_idx = 1
    while mat_col_idx < (num_cols - 2):
        for mat_line_idx in range(0,num_rows):
            if float(the_mat[mat_line_idx][prev_year_idx]) > 0.0:
                the_mat[mat_line_idx][mat_col_idx] = float(the_mat[mat_line_idx][cur_year_idx])/float(the_mat[mat_line_idx][prev_year_idx]) - 1.0
        prev_year_idx = cur_year_idx
        cur_year_idx += 2
        mat_col_idx += 2

    # fill in the 5 year change and annualized change
    cur_year_idx = num_cols - 4
    for mat_line_idx in range(0,num_rows):
        mat_col_idx = num_cols - 2
        if float(the_mat[mat_line_idx][1]) > 0.0:
            the_mat[mat_line_idx][mat_col_idx] = float(the_mat[mat_line_idx][cur_year_idx])/float(the_mat[mat_line_idx][1]) - 1.0
        mat_col_idx += 1
        x = 1.0 + float(the_mat[mat_line_idx][mat_col_idx-1])
        the_mat[mat_line_idx][mat_col_idx] =  math.pow(abs(x) ,(1.0/5.0))*(1,-1)[x<0] - 1.0
    return

def fill_section(section, year_idx, the_df, the_mat, mat_line_idx, mat_col_idx, row_label, subtotals):
    col_label = None
    ig_tot = 0.0
    hy_tot = 0.0
    for fid, class_id in zip(section, range(0,len(section))):
        if fid == '':
            mat_line_idx += 1
            continue
        col_label = get_col_label(year_idx, fid)
        # now have labels needed to lookup data in df_sch_df
        the_mat[mat_line_idx][mat_col_idx] += float(the_df.loc[row_label, col_label])
        if subtotals:
            if(class_id < 3):
                ig_tot += the_mat[mat_line_idx][mat_col_idx]
            else:
                hy_tot += the_mat[mat_line_idx][mat_col_idx]
        mat_line_idx += 1

    if subtotals:
        the_mat[mat_line_idx][mat_col_idx] = ig_tot
        mat_line_idx += 1
        the_mat[mat_line_idx][mat_col_idx] = hy_tot
        mat_line_idx += 1
        the_mat[mat_line_idx][mat_col_idx] = ig_tot + hy_tot
        mat_line_idx += 1

    return mat_line_idx

def fill_section_sum(section, year_idx, the_df, the_mat, mat_line_idx, mat_col_idx, row_label, subtotals):
    col_label = None
    ig_tot = 0.0
    hy_tot = 0.0
    for fid, class_id in zip(section, range(0, len(section))):
        if type(fid) == str:
            if fid == '':
                mat_line_idx += 1
                continue
            col_label = get_col_label(year_idx, fid)
            # now have labels needed to lookup data in df_sch_df
            the_mat[mat_line_idx][mat_col_idx] += float(the_df.loc[row_label, col_label])
            if subtotals:
                if (class_id < 2):
                    ig_tot += the_mat[mat_line_idx][mat_col_idx]
                else:
                    hy_tot += the_mat[mat_line_idx][mat_col_idx]
            mat_line_idx += 1
        elif type(fid) == list:
            for sub_fid in fid:
                col_label = get_col_label(year_idx, sub_fid)
                the_mat[mat_line_idx][mat_col_idx] += float(the_df.loc[row_label, col_label])
            pass
        else:
            logger.error("Unknown type in List: %s", row_label)

    if subtotals:
        the_mat[mat_line_idx][mat_col_idx] = ig_tot
        mat_line_idx += 1
        the_mat[mat_line_idx][mat_col_idx] = hy_tot
        mat_line_idx += 1
        the_mat[mat_line_idx][mat_col_idx] = ig_tot + hy_tot
        mat_line_idx += 1

    return mat_line_idx

def excel_print_sheet(xl_sheet, r1, r2, r3, co, column_heading, the_mat):
    # push it out to the spreadsheet
    attempt = 0
    while True:
        try:
            xl_range = xl_sheet.range(r1)
            xl_range.value = co

            xl_range = xl_sheet.range(r2)
            xl_range.value = column_heading

            xl_range = xl_sheet.range(r3)
            xl_range.value = the_mat
            break
        except:
            logger.error("Unable to write to excel workbook: %d", attempt)
            attempt += 1
            if attempt > 5:
                return
    logger.info("Leave")
    return

def excel_reset_sheets(xl_wb):
    logger.info("Enter")
    attempt = 0
    while True:
        try:
            sch_d_sheet = xl_wb.sheets(sch_d)
            sch_d_range = sch_d_sheet.range('D6:P108')
            sch_d_range.clear_contents()

            fin_sheet = xl_wb.sheets(fin)
            fin_range = fin_sheet.range('E4:Q67')
            fin_range.clear_contents()

            sch_ba_sheet = xl_wb.sheets(sch_ba)
            sch_ba_range = sch_ba_sheet.range('D4:P29')
            sch_ba_range.clear_contents()
            sch_ba_range = sch_ba_sheet.range('R4:AD29')
            sch_ba_range.clear_contents()
            sch_ba_range = sch_ba_sheet.range('AF4:AQ29')
            sch_ba_range.clear_contents()

            sg_d_sheet = xl_wb.sheets(sg_d)
            sg_d_range = sg_d_sheet.range('D6:E26')
            sg_d_range.clear_contents()
            break
        except:
            logger.error("Unable to write to excel workbook: %d", attempt)
            attempt += 1
            if attempt > 5:
                return

    logger.info("Leave")
    return

############### Build #########################################################
def build_fin(xl_sheet, co, fin_sub_tots):
    logger.info("Enter")
    row_label = co
    num_years = len(years)
    num_cols = (num_years - 1) + num_years + 2
    num_rows = len(fin_sub_tots[0]) + len(fin_sub_tots[1]) + len(fin_sub_tots[2]) + 3
    fin_mat = [[0.0 for x in range(num_cols)] for y in range(num_rows)]

    column_heading = get_column_headings(years, num_cols)

    mat_col_idx = 0
    year_idx = 0
    while year_idx < len(years):
        if mat_col_idx > 0 and (mat_col_idx % 2) == 0:
            # leave blank columns to be filled in with calculated values
            mat_col_idx += 1
            continue
        mat_line_idx = 0
        for subset in fin_sub_tots:
            mat_line_idx = fill_section_sum(subset, year_idx, fin_df, fin_mat, mat_line_idx, mat_col_idx, row_label, False)
            mat_line_idx += 1
        mat_col_idx += 1
        year_idx += 1

    calc_percent_columns(fin_mat)
    excel_print_sheet(xl_sheet, 'B1', 'E4', 'E7', co, column_heading, fin_mat)
    logger.info("Leave")
    return

def build_sch_d_vert(xl_sheet, co, accet_class_tots, combo_asset_classes):
    logger.info("Enter: %s", co)
    row_label = co
    num_years = len(years)
    num_cols = (num_years - 1) + num_years + 2
    num_rows = len(best.PC_Asset_Class_Tots) * best.LINES_PER_ASSET_CLASS + 3
    sch_d_mat = [[0.0 for x in range(num_cols)] for y in range(num_rows)]

    column_heading = get_column_headings(years, num_cols)

    mat_col_idx = 0
    year_idx = 0
    while year_idx < len(years):
        if mat_col_idx > 0 and (mat_col_idx % 2) == 0:
            # leave blank columns to be filled in with calculated values
            mat_col_idx += 1
            continue
        mat_line_idx = 0
        for asset_class in accet_class_tots:
            if asset_class in combo_asset_classes:
                tmp_mat_line_idx = mat_line_idx
                for component_asset_class in asset_class:
                    mat_line_idx = tmp_mat_line_idx
                    mat_line_idx = fill_section_sum(component_asset_class, year_idx, sch_d_df, sch_d_mat, mat_line_idx, mat_col_idx, row_label, True)
                mat_line_idx += 1
            else:
                mat_line_idx = fill_section_sum(asset_class, year_idx, sch_d_df, sch_d_mat, mat_line_idx, mat_col_idx, row_label, True)
                mat_line_idx += 1
        mat_col_idx += 1

        year_idx += 1

    calc_percent_columns(sch_d_mat)
    excel_print_sheet(xl_sheet, 'B1', 'D4', 'D6', co, column_heading, sch_d_mat)
    return

def build_sch_ba(xl_sheet, co, sch_ba_subcat):
    logger.info("Enter")
    row_label = co
    num_years = len(years)
    num_cols = (num_years - 1) + num_years + 2
    num_rows = len(sch_ba_subcat[0]) + 1
    sch_ba_mat_unaff = [[0.0 for x in range(num_cols)] for y in range(num_rows)]
    sch_ba_mat_aff = [[0.0 for x in range(num_cols)] for y in range(num_rows)]

    column_heading = get_column_headings(years, num_cols)

    mat_col_idx = 0
    year_idx = 0
    while year_idx < len(years):
        if mat_col_idx > 0 and (mat_col_idx % 2) == 0:
            # leave blank columns to be filled in with calculated values
            mat_col_idx += 1
            continue
        mat_line_idx = 0
        fill_section_sum(sch_ba_subcat[0], year_idx, sch_ba_df, sch_ba_mat_unaff, mat_line_idx, mat_col_idx, row_label, False)
        mat_line_idx = 0
        fill_section_sum(sch_ba_subcat[0], year_idx, sch_ba_df, sch_ba_mat_aff, mat_line_idx, mat_col_idx, row_label, False)
        mat_col_idx += 1
        year_idx += 1

    calc_percent_columns(sch_ba_mat_unaff)
    calc_percent_columns(sch_ba_mat_aff)
    excel_print_sheet(xl_sheet, 'B1', 'D4', 'D5', co, column_heading, sch_ba_mat_unaff)
    excel_print_sheet(xl_sheet, 'B1', 'R4', 'R5', co, column_heading, sch_ba_mat_unaff)
    logger.info("Leave")
    return

def build_sg_d(xl_sheet, co, all_single_data):
    logger.info("Enter")
    row_label = co
    num_years = len(years)
    num_cols = 2
    num_rows = len(best.ALL_Single_Data) + 1
    sg_d_mat = [["" for x in range(num_cols)] for y in range(num_rows)]
    # fill_section will not work for this sheet - no years and multiple columns  required

    column_heading = ["Col 1", "Col 2"]
    mat_line_idx = 0
    for fid in all_single_data:
        if type(fid) == str:
            if fid == '':
                mat_line_idx += 1
                continue
        col_label = get_col_label(0, fid)
        sg_d_mat[mat_line_idx][0] = ""
        sg_d_mat[mat_line_idx][1] =  str(sg_d_df.loc[row_label, col_label])
        if col_label == "CO0010":
            sg_d_mat[mat_line_idx][0] = str(sg_d_df.loc[row_label, "CO00011"])
        if col_label == "CO00231":
            company_map[row_label] = str(sg_d_df.loc[row_label, "CO00231"])
        mat_line_idx += 1

    excel_print_sheet(xl_sheet, 'B1', 'D5', 'D6', co, column_heading, sg_d_mat)
    logger.info("Leave")
    return

def build_4_sheets(xl_wb, entity_id):
    logger.info("Enter")

    excel_reset_sheets(xl_wb)
    company_count = 0
    fin_sheet = xl_wb.sheets(fin)
    sch_d_sheet = xl_wb.sheets(sch_d)
    sch_ba_sheet = xl_wb.sheets(sch_ba)
    sg_d_sheet = xl_wb.sheets(sg_d)
    for co in companies:
        attempt = 0
        if entity_id == 'PC':
            build_sch_d_vert(sch_d_sheet, co, best.PC_Asset_Class_Tots, best.PC_Combo_Asset_Classes)
            build_sch_ba(sch_ba_sheet, co, best.PC_Subcat)
            build_fin(fin_sheet, co, best.PC_Fin_Sub_Tots)
            build_sg_d(sg_d_sheet, co, best.ALL_Single_Data)
        if entity_id == 'LIFE':
            build_sch_d_vert(sch_d_sheet, co, best.LIFE_Asset_Class_Tots, best.LIFE_Combo_Asset_Classes)
            build_sch_ba(sch_ba_sheet, co, best.LIFE_Subcat)
            build_fin(fin_sheet, co, best.LIFE_Fin_Sub_Tots)
            pass
        if entity_id == 'HEALTH':
            build_sch_d_vert(sch_d_sheet, co, best.HEALTH_Asset_Class_Tots, best.HEALTH_Combo_Asset_Classes)
            build_sch_ba(sch_ba_sheet, co, best.HEALTH_Subcat)
            build_fin(fin_sheet, co, best.HEALTH_Fin_Sub_Tots)
            pass
        while True:
            try:
                xl_mac = xl_wb.macro('CalcSheet')
                xl_mac(ts, company_map[co])
                #        ts_sheet = xl_wb.sheets(ts)
                #        tmp_sheet = xl_wb.sheets(co)
                #        tmp_sheet.range('A1').value = ts_sheet.cells.value
                break
            except:
                logger.error("Unable to write to excel workbook: %d", attempt)
                attempt += 1
                if attempt > 5:
                    return
        company_count += 1
    logger.info("Leave: company_count = %d", company_count)
    return

# excel entry point
def create_tearsheets():
    logger.info("Enter")
    global companies

#    xl_wb = xw.Book.caller()
    xl_wb = xw.Book(r'C:\Users\tjevon\Dropbox\code\sandbox\aumi\xl\TearSheet_py2.xlsm')

    # findout what sheets are already open in the workbook
    available_sheets = []
    for i in xl_wb.sheets:
        logger.debug(i.name)
        available_sheets.append(i.name)
    logger.debug("Have Sheets:")

    for sheet_name in needed_sheets:
        if sheet_name not in available_sheets:
            xl_wb.sheets.add(sheet_name)

    # get a dictionary of the files {PC:[sch_d, sch_ba, fin], LIFE:[sch_d, sch_ba_fin], etc...}
    file_info = get_file_dict()

    files_with_all = "ALL"
    read_csv_all(xl_wb, files_with_all, file_info[files_with_all])

    # for each grouping, read the csv's into a DataFrame and build the 4 xl_sheets
    for entity_id in file_info:
        companies = set()
        logger.info("Industry: %s", entity_id)
        read_csvs(xl_wb, entity_id, file_info[entity_id])
#        for sheet_name in companies:
#            if sheet_name not in available_sheets:
#                xl_wb.sheets.add(sheet_name)
        build_4_sheets(xl_wb, entity_id)
    logger.info("Leave")
    return

################################################################################
if __name__ == '__main__':
    create_tearsheets()
