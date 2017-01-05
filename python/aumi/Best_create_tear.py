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
companies = set()

needed_sheets = [sch_d, sch_ba, fin, sg_d]

LIFE_Unaffiliated = ['ST13332', 'ST13334', 'ST13336', 'ST13338', 'ST13340', 'ST13342', 'ST13344', 'ST13346',
                     'ST13348', 'ST25939', 'ST13350', 'ST13352', 'ST13354', 'ST16031', 'ST13356', 'ST15400',
                     'ST15402', 'ST25941', 'ST25943', '', 'ST15406', 'ST25945', 'ST13360', '']

LIFE_Affiliated =   ['ST13333', 'ST13335', 'ST13337', 'ST13339', 'ST13341', 'ST13343', 'ST13345', 'ST13347',
                     'ST13349', 'ST25940', 'ST13351', 'ST13353', 'ST13355', 'ST16032', 'ST13357', 'ST15401',
                     'ST15403', 'ST25942', 'ST25944', '', 'ST15407', '', 'ST13361', '']

LIFE_Subcat = [LIFE_Unaffiliated, LIFE_Affiliated]

PC_Net_Pos =     ['SF13840', 'SF00011', '', 'SF00013', 'SF00014', '', 'SF04762', 'SF04763', 'SF00016', 'SF06017', 'SF06018', 'SF04764', 'SF00023', 'SF00025','', 'SF00031', '', '', '']
PC_Iris_Ratios =     ['SF05352', 'SF05354', 'SF05356', 'SF05358', 'SF05360', 'SF05362', 'SF08471', 'SF08473', 'SF05366', 'SF05368', 'SF05370', 'SF05372', 'SF05374','','','','','','','','','','','','']
PC_Liquidity =   ['','CalcS80050', '','','CalcS80031', 'SF00025', 'SF00031', 'SF08536', 'SF00053', 'SF00139', 'SF07654', 'SF07656', 'SF00148', 'SF00147', 'SF00146']

PC_Fin_Sub_Tots = [PC_Iris_Ratios, PC_Liquidity, PC_Net_Pos]

LIFE_Iris_Ratios =   ['','','','','','','','','','','','','','ST08607', 'ST08609', 'ST08611', 'ST08613', 'ST08615', 'ST08617', 'ST08619', 'ST08621', 'ST08623', 'ST08625', 'ST08627', 'ST08629']


PC_US_Fed_Tot =        ['SF04565', 'SF04566', 'SF04567', 'SF04568', 'SF04569', 'SF04570']
PC_Oth_Gov_Tot =       ['SF10297', 'SF10298', 'SF10299', 'SF10300', 'SF10301', 'SF10302']
PC_US_St_Ter_Tot =     ['SF10304', 'SF10305', 'SF10306', 'SF10307', 'SF10308', 'SF10309']
PC_US_Pol_Sub_Tot =    ['SF10311', 'SF10312', 'SF10313', 'SF10314', 'SF10315', 'SF10316']
PC_US_Spec_Rev_Tot =   ['SF10318', 'SF10319', 'SF10320', 'SF10321', 'SF10322', 'SF10323']
PC_Corp_Tot =          ['SF10325', 'SF10326', 'SF10327', 'SF10328', 'SF10329', 'SF10330']
PC_Hybrid_Tot =        ['SF10332', 'SF10333', 'SF10334', 'SF10335', 'SF10336', 'SF10337']
PC_Par_Sub_Aff_Tot =   ['SF02189', 'SF02190', 'SF02191', 'SF02192', 'SF02193', 'SF02194']
PC_Gov_Ag_Tot =        [PC_US_Fed_Tot, PC_Oth_Gov_Tot]
PC_Muni_Tot =          [PC_US_St_Ter_Tot, PC_US_Pol_Sub_Tot, PC_US_Spec_Rev_Tot]

PC_Combo_Asset_Classes = [PC_Gov_Ag_Tot, PC_Muni_Tot]
PC_Asset_Class_Tots = [PC_US_Fed_Tot, PC_Oth_Gov_Tot, PC_US_St_Ter_Tot, PC_US_Pol_Sub_Tot, PC_Gov_Ag_Tot, PC_Corp_Tot, PC_US_Spec_Rev_Tot, PC_Hybrid_Tot, PC_Par_Sub_Aff_Tot, PC_Muni_Tot]

LIFE_US_Fed_Tot =        ['ST07163', 'ST07164', 'ST07165', 'ST07166', 'ST07167','ST07168']
LIFE_Oth_Gov_Tot =       ['ST18057', 'ST18058', 'ST18059', 'ST18060', 'ST18061', 'ST18062']
LIFE_US_St_Ter_Tot =     ['ST18064', 'ST18065', 'ST18066', 'ST18067', 'ST18068', 'ST18069']
LIFE_US_Pol_Sub_Tot =    ['ST18071', 'ST18072', 'ST18073', 'ST18074', 'ST18075', 'ST18076']
LIFE_US_Spec_Rev_Tot =   ['ST18078', 'ST18079', 'ST18080', 'ST18081', 'ST18082', 'ST18083']
LIFE_Corp_Tot =          ['ST18085', 'ST18086', 'ST18087', 'ST18088', 'ST18089', 'ST18090']
LIFE_Hybrid_Tot =        ['ST18092', 'ST18093', 'ST18094', 'ST18095', 'ST18096', 'ST18097']
LIFE_Par_Sub_Aff_Tot =   ['ST04422', 'ST04423', 'ST04424', 'ST04425', 'ST04426', 'ST04427']
LIFE_Gov_Ag_Tot =        [LIFE_US_Fed_Tot, LIFE_Oth_Gov_Tot]
LIFE_Muni_Tot =          [LIFE_US_St_Ter_Tot, LIFE_US_Pol_Sub_Tot, LIFE_US_Spec_Rev_Tot]

LIFE_Combo_Asset_Classes = [LIFE_Gov_Ag_Tot, LIFE_Muni_Tot]
LIFE_Asset_Class_Tots = [LIFE_US_Fed_Tot, LIFE_Oth_Gov_Tot, LIFE_US_St_Ter_Tot, LIFE_US_Pol_Sub_Tot, LIFE_Gov_Ag_Tot, LIFE_Corp_Tot, LIFE_US_Spec_Rev_Tot, LIFE_Hybrid_Tot, LIFE_Par_Sub_Aff_Tot, LIFE_Muni_Tot]

HEALTH_US_Fed_Tot =        ['HS02710', 'HS02711', 'HS02712', 'HS02713', 'HS02714', 'HS02715']
HEALTH_Oth_Gov_Tot =       ['HS10291', 'HS10292', 'HS10293', 'HS10294', 'HS10295', 'HS10296']
HEALTH_US_St_Ter_Tot =     ['HS10298', 'HS10299', 'HS10300', 'HS10301', 'HS10302', 'HS10303']
HEALTH_US_Pol_Sub_Tot =    ['HS10305', 'HS10306', 'HS10307', 'HS10308', 'HS10309', 'HS10310']
HEALTH_US_Spec_Rev_Tot =   ['HS10312', 'HS10313', 'HS10314', 'HS10315', 'HS10316', 'HS10317']
HEALTH_Corp_Tot =          ['HS10319', 'HS10320', 'HS10321', 'HS10322', 'HS10323', 'HS10324']
HEALTH_Hybrid_Tot =        ['HS10326', 'HS10327', 'HS10328', 'HS10329', 'HS10330', 'HS10331']
HEALTH_Par_Sub_Aff_Tot =   ['HS03011', 'HS03012', 'HS03013', 'HS03014', 'HS03015', 'HS03016']
HEALTH_Gov_Ag_Tot =        [HEALTH_US_Fed_Tot, HEALTH_Oth_Gov_Tot]
HEALTH_Muni_Tot =          [HEALTH_US_St_Ter_Tot, HEALTH_US_Pol_Sub_Tot, HEALTH_US_Spec_Rev_Tot]

HEALTH_Combo_Asset_Classes = [HEALTH_Gov_Ag_Tot, HEALTH_Muni_Tot]
HEALTH_Asset_Class_Tots = [HEALTH_US_Fed_Tot, HEALTH_Oth_Gov_Tot, HEALTH_US_St_Ter_Tot, HEALTH_US_Pol_Sub_Tot, HEALTH_Gov_Ag_Tot, HEALTH_Corp_Tot, HEALTH_US_Spec_Rev_Tot, HEALTH_Hybrid_Tot, HEALTH_Par_Sub_Aff_Tot, HEALTH_Muni_Tot]

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

def read_sch_d(entity_id, csv_filename):
    logger.info("Enter")
    global sch_d_df
    global years
    global companies
    years = []
    tmp_years = []
    sch_d_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode')
    for i in sch_d_df.iloc[1]:
        if pd.notnull(i):
            tmp_years.append(int(i))
    years = sorted(set(tmp_years))
    # should read df with group # for column lable and field # for row label
    sch_d_df = pd.read_csv(csv_filename, header=3, index_col=0, dtype='unicode')
    for i in sch_d_df.columns:
        if i.find("Unnamed") != -1:
            continue
        if i.find(".") != -1:
            continue
        if pd.notnull(i):
            companies.add(i)
    pass
    logger.info("Leave")

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
#    companies = sorted[set(companies)]
    for i in the_df.columns:
        if i.find('Calc') != -1:
            the_df[i].replace(regex=True,inplace=True,to_replace=r'%',value=r'')

    logger.info("Leave")
    return the_df

def read_csvs( xl_wb, entity_id, file_names ):
    logger.info("Enter")
    global sch_ba_df
    global sch_d_df
    global fin_df
    global sg_d_df

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

############### Build #########################################################
def build_4_sheets(xl_wb, entity_id):
    logger.info("Enter")
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
    sch_ba_range = sch_ba_sheet.range('AF4:AR29')
    sch_ba_range.clear_contents()

    for co in companies:
#        if entity_id == 'PC':
#            build_sch_d_vert(sch_d_sheet, co, PC_Asset_Class_Tots, PC_Combo_Asset_Classes)
#            build_fin(fin_sheet, co, PC_Fin_Sub_Tots)
        if entity_id == 'LIFE':
#            build_sch_d_vert(sch_d_sheet, co, LIFE_Asset_Class_Tots, LIFE_Combo_Asset_Classes)
            build_sch_ba(sch_ba_sheet, co, LIFE_Subcat)
#        if entity_id == 'HEALTH':
#            build_sch_d_vert(sch_d_sheet, co, HEALTH_Asset_Class_Tots, HEALTH_Combo_Asset_Classes)
    #    build_fin(fin_sheet, co)
    logger.info("Leave")

def build_fin(xl_sheet, co, fin_sub_tots):
    logger.info("Enter")

    # first cut at vertical file orientation, only looking at a single row in csv at a time
    row_label = co

    # TODO: break this out
    num_years = len(years)
    num_cols = (num_years - 1) + num_years + 2

    num_rows = len(fin_sub_tots[0]) + len(fin_sub_tots[1]) + len(fin_sub_tots[2]) + 3

    # allocate 2 D list for numeric data
    fin_mat = [[0.0 for x in range(num_cols)] for y in range(num_rows)]

    column_heading = []
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
            mat_line_idx = fill_section(subset, year_idx, fin_df, fin_mat, mat_line_idx, mat_col_idx, row_label, False)
            mat_line_idx += 1
        mat_col_idx += 1

        year_idx += 1

    # fill in the blank columns with % change data
    # fill in the year over year change
    mat_col_idx = 2
    prev_year_idx = 0
    cur_year_idx = 1
    while mat_col_idx < (num_cols - 2):
        for mat_line_idx in range(0,num_rows):
            if float(fin_mat[mat_line_idx][prev_year_idx]) > 0.0:
                fin_mat[mat_line_idx][mat_col_idx] = float(fin_mat[mat_line_idx][cur_year_idx])/float(fin_mat[mat_line_idx][prev_year_idx]) - 1.0
        prev_year_idx = cur_year_idx
        cur_year_idx += 2
        mat_col_idx += 2

    # fill in the 5 year change and annualized change
    cur_year_idx = num_cols - 4
    for mat_line_idx in range(0,num_rows):
        mat_col_idx = num_cols - 2
        if float(fin_mat[mat_line_idx][1]) > 0.0:
            fin_mat[mat_line_idx][mat_col_idx] = float(fin_mat[mat_line_idx][cur_year_idx])/float(fin_mat[mat_line_idx][1]) - 1.0
        mat_col_idx += 1
#        fin_mat[mat_line_idx][mat_col_idx] = (1.0 + float(fin_mat[mat_line_idx][mat_col_idx-1])) ** (1.0/5.0) - 1.0
    pass

    # push it out to the spreadsheet
    attempt = 0
    while True:
        try:
            xl_range = xl_sheet.range('B1')
            xl_range.value = co

            xl_range = xl_sheet.range('E4')
            xl_range.value = column_heading

            xl_range = xl_sheet.range('E7')
            xl_range.value = fin_mat
            break
        except:
            logger.error("Unable to write to excel workbook: %d", attempt)
            attempt += 1
            if attempt > 5:
                return
    logger.info("Leave")

def build_sch_d_vert(xl_sheet, co, accet_class_tots, combo_asset_classes):

    logger.info("Enter: %s", co)

    row_label = co

    num_years = len(years)
    num_cols = (num_years - 1) + num_years + 2
    num_rows = len(PC_Asset_Class_Tots) * LINES_PER_ASSET_CLASS + 3

    # allocate 2 D list for numeric data
    sch_d_mat = [[0.0 for x in range(num_cols)] for y in range(num_rows)]

    column_heading = []
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
                    mat_line_idx = fill_section(component_asset_class, year_idx, sch_d_df, sch_d_mat, mat_line_idx, mat_col_idx, row_label, True)
                mat_line_idx += 1
            else:
                mat_line_idx = fill_section(asset_class, year_idx, sch_d_df, sch_d_mat, mat_line_idx, mat_col_idx, row_label, True)
                mat_line_idx += 1
        mat_col_idx += 1

        year_idx += 1

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
        sch_d_mat[mat_line_idx][mat_col_idx] = (1.0 + float(sch_d_mat[mat_line_idx][mat_col_idx-1])) ** (1.0/5.0) - 1.0
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
    logger.info("Leave")

def build_sch_ba(xl_sheet, co, sch_ba_subcat):
    logger.info("Enter")

    # first cut at vertical file orientation, only looking at a single row in csv at a time
    row_label = co

    # TODO: break this out
    num_years = len(years)
    num_cols = (num_years - 1) + num_years + 2

    num_rows = len(sch_ba_subcat[0]) + len(sch_ba_subcat[1]) + 3

    # allocate 2 D list for numeric data
    sch_ba_mat = [[0.0 for x in range(num_cols)] for y in range(num_rows)]

    column_heading = []
    column_heading = get_column_headings(years, num_cols)

    mat_col_idx = 0
    year_idx = 0
    while year_idx < len(years):
        if mat_col_idx > 0 and (mat_col_idx % 2) == 0:
            # leave blank columns to be filled in with calculated values
            mat_col_idx += 1
            continue
        mat_line_idx = 0
        for subset in sch_ba_subcat:
            mat_line_idx = fill_section(subset, year_idx, sch_ba_df, sch_ba_mat, mat_line_idx, mat_col_idx, row_label, False)
            mat_line_idx += 1
        mat_col_idx += 1

        year_idx += 1

   # fill in the blank columns with % change data
    # fill in the year over year change
    mat_col_idx = 2
    prev_year_idx = 0
    cur_year_idx = 1
    while mat_col_idx < (num_cols - 2):
        for mat_line_idx in range(0,num_rows):
            if float(sch_ba_mat[mat_line_idx][prev_year_idx]) > 0.0:
                sch_ba_mat[mat_line_idx][mat_col_idx] = float(sch_ba_mat[mat_line_idx][cur_year_idx])/float(sch_ba_mat[mat_line_idx][prev_year_idx]) - 1.0
        prev_year_idx = cur_year_idx
        cur_year_idx += 2
        mat_col_idx += 2

    # fill in the 5 year change and annualized change
    cur_year_idx = num_cols - 4
    for mat_line_idx in range(0,num_rows):
        mat_col_idx = num_cols - 2
        if float(sch_ba_mat[mat_line_idx][1]) > 0.0:
            sch_ba_mat[mat_line_idx][mat_col_idx] = float(sch_ba_mat[mat_line_idx][cur_year_idx])/float(sch_ba_mat[mat_line_idx][1]) - 1.0
        mat_col_idx += 1
#        sch_ba_mat[mat_line_idx][mat_col_idx] = (1.0 + float(sch_ba_mat[mat_line_idx][mat_col_idx-1])) ** (1.0/5.0) - 1.0
    pass

    # push it out to the spreadsheet
    attempt = 0
    while True:
        try:
            xl_range = xl_sheet.range('B1')
            xl_range.value = co

            xl_range = xl_sheet.range('E4')
            xl_range.value = column_heading

            xl_range = xl_sheet.range('E7')
            xl_range.value = sch_ba_mat
            break
        except:
            logger.error("Unable to write to excel workbook: %d", attempt)
            attempt += 1
            if attempt > 5:
                return
    logger.info("Leave")


def build_sch_d(xl_sheet, co):
    logger.info("Enter")
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
        sch_d_mat[mat_line_idx][mat_col_idx] = (1.0 + float(sch_d_mat[mat_line_idx][mat_col_idx-1])) ** (1.0/5.0) - 1.0
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
    logger.info("Leave")

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

#Entry point
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
        xl_sheet = xl_wb.sheets(sheet_name)

    file_info = get_file_dict()

    # for each company, read in the csv's and calculate the tear sheet
    for entity_id in file_info:
        companies = set()
        logger.info("Industry: %s", entity_id)
        read_csvs(xl_wb, entity_id, file_info[entity_id])
        if entity_id == "LIFE":
            build_4_sheets(xl_wb, entity_id)
#        xl_mac = xl_wb.macro('CalcSheet')
#        xl_mac(ts)
    logger.info("Leave")

################################################################################
if __name__ == '__main__':
    create_tearsheets()
