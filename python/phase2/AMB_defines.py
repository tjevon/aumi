import pandas as pd
import numpy as np
import logging
import os
import locale

logger = logging.getLogger('twolane')

BAD_CHAR = ["/", ":", "*", "[", "]", ":", "?"]

# tag - used to locate filenames and template worksheets
SI01_tag     = "SI01"
E07_tag      = "E07"
E10_tag      = "E10"
Liab1_tag    = "Liab1"
Liab2_tag    = "Liab2"
Liab3_tag    = "Liab3"
SI05_07_tag  = "SI05_07"
SI08_09_tag  = "SI08_09"
Assets_tag   = "Assets"
SoI_tag      = "SoI"
SoO_tag      = "SoO"
SoR_tag      = "SoR"
IRIS1_tag    = "IRIS1"
IRIS2_tag    = "IRIS2"
CashFlow_tag = "CashFlow"
CR_tag       = "CR"

PC_tag      = "PC"
LIFE_tag    = "LIFE"
HEALTH_tag  = "HEALTH"

BUSINESS_TYPES = [PC_tag, LIFE_tag, HEALTH_tag]
YEARLY_IDX = 0
QUARTERLY_IDX = 1
DISPLAY_IDX = 2
MPL_IDX = 3
CALC_COL = 'H'

COMMON_TEMPLATE_TAGS = [Assets_tag, E10_tag, E07_tag, CashFlow_tag, SI01_tag, SI05_07_tag]
PC_TEMPLATE_TAGS = [SoI_tag]
LIFE_TEMPLATE_TAGS = [SoO_tag]
HEALTH_TEMPLATE_TAGS = [SoR_tag]
#PC_TEMPLATE_TAGS = [SoI_tag, IRIS1_tag ]
#LIFE_TEMPLATE_TAGS = [SoO_tag, IRIS2_tag]

COMMON_QUARTERLY_TAGS = [Assets_tag, CashFlow_tag]
PC_QUARTERLY_TAGS = [SoI_tag]
LIFE_QUARTERLY_TAGS = [SoO_tag]
HEALTH_QUARTERLY_TAGS = [SoR_tag]

BUSINESS_TYPE_TAGS = {YEARLY_IDX:{PC_tag: COMMON_TEMPLATE_TAGS+PC_TEMPLATE_TAGS,
                      LIFE_tag: COMMON_TEMPLATE_TAGS+LIFE_TEMPLATE_TAGS,
                      HEALTH_tag: COMMON_TEMPLATE_TAGS+HEALTH_TEMPLATE_TAGS},
                      QUARTERLY_IDX:{PC_tag: PC_QUARTERLY_TAGS+COMMON_QUARTERLY_TAGS,
                                     LIFE_tag: LIFE_QUARTERLY_TAGS+COMMON_QUARTERLY_TAGS,
                                     HEALTH_tag: HEALTH_QUARTERLY_TAGS+COMMON_QUARTERLY_TAGS}}

TS_TEMPLATES = {PC_tag:'PC_Template', LIFE_tag:'LIFE_template', HEALTH_tag:'HEALTH_template'}

PERCENT_FORMATS = [IRIS1_tag, IRIS2_tag]

DATA_DIR = os.getenv('DATAPATH', './')

#YEARLY_DIR = "AAA"
#QTRLY_DIR = "AAA_Q"
#COMPANY_MAP_DIR = "AAA_company_mapping"
#COMPANY_INFO_DIR = "AAA_company_info"

#YEARLY_DIR = "test_data"
#QTRLY_DIR = "test_qtrly"
#YEARLY_DIR = "tmp"
#QTRLY_DIR = "tmp_Q"
#COMPANY_INFO_DIR = "tmp_company_info"
#COMPANY_MAP_DIR = "test_mapping"
#COMPANY_INFO_DIR = "test_company_info"
#COMPANY_INFO_DIR = "2013_2016_5_company_info"
#POSITIONS_DIR = "2015_2016_investments"

#YEARLY_DIR = "allstate_data"
#QTRLY_DIR = "allstate_qtrly"
#COMPANY_MAP_DIR = "allstate_mapping"

#YEARLY_DIR = "2006_2016"
#QTRLY_DIR = "2016_qtrly"

YEARLY_DIR = "2013_2016_5_data"
QTRLY_DIR = "2013_2016_5Q_data"
COMPANY_INFO_DIR = "2013_2016_5_company_info"
COMPANY_MAP_DIR = "2013_2016_5_company_mapping"

#COMPANY_INFO_DIR = "company_info"
#COMPANY_MAP_DIR = "company_mapping"
POSITIONS_DIR = "inv_unaf_af_0"

TEMPLATE_DIR = "templates"
OUTPUT_DIR = "output"


TEMPLATE_FILENAME = "TearSheet_Template.xlsx"

TARGET_FILENAME = "TearSheet_Output.xlsx"
TARGET_SHEET = "TS"

DO_NOT_DISPLAY =        0x00

DISPLAY_SECTION =       0x01
DISPLAY_PROJECTION =    0x02
CALCULATE_PROJECTION =  0x04

def do_all_calculations(template_wb, section_map, data_cube, position_cube, desired_periods):
    ai_info = []
    bi_info = []
    ci_info = []
    vi_info = []
    zi_info = []
    for tag, section in section_map.iteritems():
        ai_info.append((tag, section.comp_dict_ai, section.fid_collection_dict))
        bi_info.append((tag, section.comp_dict_bi, section.fid_collection_dict))
        ci_info.append((tag, section.comp_dict_ci, section.fid_collection_dict))
        vi_info.append((tag, section.comp_dict_vi, section.fid_collection_dict))
        zi_info.append((tag, section.comp_dict_zi, section.fid_collection_dict))
    dependency_ordered_list = [vi_info, zi_info, ai_info, bi_info, ci_info]

    for calculation_level in dependency_ordered_list:
        for calc in calculation_level:
            data_cube = do_calculation(calc, template_wb, data_cube, section_map, position_cube, desired_periods)
    return data_cube


def do_calculation(calc, template_wb, data_cube, section_map, position_cube, desired_periods):
    df_dict = {}
    tag = calc[0]
    comp_dict = calc[1]
    giveup = False
    fid_collection_dict = calc[2]
    slice1 = None
    for key, fids in comp_dict.iteritems():
        cell = key.replace('A', CALC_COL)
        formula = template_wb.get_formula(tag, cell)
        formula = formula.replace('$','')
        func_idx = formula.find('(')
        args_idx = formula.rfind(')')
        func_key = formula[:func_idx]
        args_str = formula[func_idx + 1:args_idx]
        args = args_str.split(",")
        if func_key == AI_SET:
            slice1 = func_dict[func_key](args,data_cube)
        else:
            cell_info = args.pop(0)
            fid1 = lookup_fid(section_map, fid_collection_dict, cell_info)
            if fid1 is None:
                continue

            slice1 = data_cube[fid1]
            if func_key in POSITION_CALCS:
                slice1 = func_dict[func_key](tag, args, position_cube, slice1, desired_periods)
            elif func_key == AI_2YR_AVE:
                slice1 = func_dict[func_key](slice1)
            elif func_key == AI_SET_DF:
                slice1 = func_dict[func_key](slice1)
            else:
                for row in args[0:]:
                    fid2 = lookup_fid(section_map, fid_collection_dict, row)
                    if fid2 is None:
                        giveup = True
                        break
                    slice2 = data_cube[fid2]
                    slice1 = func_dict[func_key](slice1, slice2)
                if giveup:
                    giveup = False
                    continue
        df_dict[fids[0]] = slice1

    tmp_cube = pd.Panel(df_dict)
    cube_list = [data_cube, tmp_cube]
    data_cube = pd.concat(cube_list, axis=0)
    return data_cube


def lookup_fid(section_map, fid_collection_dict, arg):
    sht_idx = arg.find('!')
    fid = None
    if sht_idx == -1:
        fid = fid_collection_dict[arg][0]
    else:
        tag = arg[:sht_idx]
        tag = tag.replace("'", "")
        tmp_arg = arg[sht_idx+1:]
        if tag in section_map:
            section_fid_collection_dict = section_map[tag].fid_collection_dict
            fid = section_fid_collection_dict[tmp_arg][0]
    return fid

def slice_set(args,cube):
    tmp_1 = cube.axes[1]
    tmp_2 = cube.axes[2]
    rv_slice = pd.DataFrame(float(args[0]), index=tmp_1, columns=tmp_2)
    return rv_slice

def two_year_ave(slice1):
    df = slice1
    df = df[df.columns[::-1]]  # reverse
    df = df.rolling(2,axis=1)
    df = df.mean()
    rv_slice = df[df.columns[::-1]]  # reverse

    return rv_slice

def slice_sum(slice1, slice2):
    rv_slice = slice1 + slice2
    return rv_slice

def slice_set_df(slice1):
    rv_slice = slice1
    return rv_slice


def slice_diff(slice1, slice2):
    rv_slice = slice1 - slice2
    return rv_slice


def slice_div(slice1, slice2):
    rv_slice = slice1 / slice2
    return rv_slice


def slice_mult(slice1, slice2):
    rv_slice = slice1 * slice2
    return rv_slice

def bond_sum_help(tag, args, position_cube, desired_periods, naic_levels, region):
    locale.setlocale(locale.LC_NUMERIC,'')
    df_list = []
    for period in desired_periods:
        if period in position_cube:
            bond_dict = position_cube[period][0]
            keys = list(bond_dict.keys())
            sum_list = []
            period_dict = {}
            for grp, df in bond_dict.iteritems():
                tmp_df = df
                if region == 'US':
                    tmp_df = tmp_df.loc[df['IB00031'] == '...']
                elif region == 'OTHER':
                    tmp_df = tmp_df.loc[df['IB00031'] != '...']
                tmp_df = tmp_df.loc[df['IB00049'].isin(args)]
                tmp_df['naic_level'] = tmp_df['IB00033'].apply(lambda x: x[:1])
                tmp_df = tmp_df.loc[tmp_df['naic_level'].isin(naic_levels)]
                tmp_df['IB00034'] = tmp_df['IB00034'].apply(locale.atof)
                sum = tmp_df['IB00034'].sum()/1000
                sum_list.append(sum)
            period_dict[period] = sum_list
            df = pd.DataFrame.from_dict(period_dict)
            df.index = keys
            df_list.append(df)
        else:
            tmp_dict = {period:[0]}
            df = pd.DataFrame.from_dict(tmp_dict)
            df.index = ['Empty']
            df_list.append(df)

    df = pd.concat(df_list, axis=1)
    if 'Empty' in df.index:
        df = df.drop('Empty')
    df.fillna(0,inplace=True)
    df = df.astype(long)
    return df

def bond_reflection(slice1, slice2):
    slice2 = slice2.loc[slice1.index]
    slice3 = slice2.div(slice1,axis=0)
    slice3 = slice3.iloc[:,0:2].mean(axis=1)
    slice3 = slice1.iloc[:,2:4].mul(slice3,axis=0)
    slice3 = pd.concat([slice2.iloc[:,0:2], slice3], axis=1)
    return slice3

def ig_sum(tag, args, position_cube, slice1, desired_periods):
    region = args.pop(0)
    slice2 = bond_sum_help(tag,args,position_cube, desired_periods,['1','2'], region)
    slice3 = bond_reflection(slice1, slice2)
    return slice3

def hy_sum(tag, args, position_cube, slice1, desired_periods):
    region = args.pop(0)
    slice2 = bond_sum_help(tag,args,position_cube, desired_periods,['3','4','5','6'], region)
    slice3 = bond_reflection(slice1, slice2)
    return slice3

#def equity_sum(tag, args, position_cube, data_cube, desired_periods):
#    strings = []
#    return equity_sum_help(tag, args, position_cube, data_cube, desired_periods, strings)

def equity_etf_sum(tag, args, position_cube, slice1, desired_periods):
    strings = [' ETF ', 'SPDR', 'PROSHARES', 'POWERSHARES', 'Exchange Traded F', 'EXCHANGE_TRADED F', 'ISHARE']
    slice2 = equity_sum_help(tag, args, position_cube, desired_periods, strings)
    slice3 = bond_reflection(slice1, slice2)
    return slice3

def equity_sum_help(tag, args, position_cube, desired_periods, strings):
    locale.setlocale(locale.LC_NUMERIC,'')
    df_list = []
    for period in desired_periods:
        if period in position_cube:
            stock_dict = position_cube[period][1]
            keys = list(stock_dict.keys())
            sum_list = []
            period_dict = {}
            for grp, df in stock_dict.iteritems():
                tmp_df = df
                tmp_df = tmp_df.loc[df['IS00049'].isin(args)]
                if len(strings) > 0:
                    tmp_df = tmp_df[(tmp_df['IS00018'].apply(lambda x: any(y in x for y in strings))) |
                                     (tmp_df['IS00021'].apply(lambda x: any(y in x for y in strings))) ]
                if tmp_df.size == 0:
                    sum = 0
                else:
                    tmp_df['IS00037'] = tmp_df['IS00037'].apply(locale.atof)
                    sum = tmp_df['IS00037'].sum()/1000
                sum_list.append(sum)
            period_dict[period] = sum_list
            df = pd.DataFrame.from_dict(period_dict)
            df.index = keys
            df_list.append(df)
        else:
            tmp_dict = {period:[0]}
            df = pd.DataFrame.from_dict(tmp_dict)
            df.index = ['Empty']
            df_list.append(df)

    df = pd.concat(df_list, axis=1)
    if 'Empty' in df.index:
        df = df.drop('Empty')
    df.fillna(0,inplace=True)
    df = df.astype(long)
    return df

AI_SUM = "=AI_SUM"
AI_DIFF = "=AI_DIFF"
AI_MULT = "=AI_MULT"
AI_DIV = "=AI_DIV"

AI_HY = "=AI_HY"
AI_IG = "=AI_IG"

AI_EQ = "=AI_EQ"
AI_EQ_ETF = "=AI_EQ_ETF"

AI_SET_DF = "=AI_SET_DF"
AI_SET = "=AI_SET"
AI_2YR_AVE = "=AI_2YR_AVE"

POSITION_CALCS = [AI_EQ_ETF, AI_HY, AI_IG]

func_dict = {
    AI_SUM: slice_sum,
    AI_DIFF: slice_diff,
    AI_MULT: slice_mult,
    AI_DIV: slice_div,
    AI_IG: ig_sum,
    AI_HY: hy_sum,
#    AI_EQ: equity_sum,
    AI_EQ_ETF: equity_etf_sum,
    AI_SET: slice_set,
    AI_SET_DF: slice_set_df,
    AI_2YR_AVE: two_year_ave
     }

class Section(object):
    def __init__(self, tag, fid_collection_dict, comp_dict_ai, comp_dict_bi, comp_dict_ci,
                 comp_dict_vi, comp_dict_zi):
        self.tag = tag
        self.fid_collection_dict = fid_collection_dict
        self.comp_dict_ai = comp_dict_ai
        self.comp_dict_bi = comp_dict_bi
        self.comp_dict_ci = comp_dict_ci
        self.comp_dict_vi = comp_dict_vi
        self.comp_dict_zi = comp_dict_zi
    pass
