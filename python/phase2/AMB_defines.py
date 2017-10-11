import pandas as pd
import logging

import locale
from Fid_defines import *
from Tag_defines import *
from Dir_defines import *

logger = logging.getLogger('twolane')

BAD_CHAR = ["/", ":", "*", "[", "]", ":", "?"]

ETF_STRINGS = [' ETF ', 'SPDR', 'PROSHARES', 'POWERSHARES', 'Exchange Traded F', 'EXCHANGE_TRADED F', 'ISHARE']

YEARLY_IDX = 0
QUARTERLY_IDX = 1
DISPLAY_IDX = 2
MPL_IDX = 3
IF_IDX = 4
CALC_COL = 'H'
CALC_Q_COL = 'I'

BUSINESS_TYPE_TAGS = {YEARLY_IDX: {PC_tag: COMMON_TEMPLATE_TAGS+PC_TEMPLATE_TAGS,
                                   LIFE_tag: COMMON_TEMPLATE_TAGS+LIFE_TEMPLATE_TAGS,
                                   HEALTH_tag: COMMON_TEMPLATE_TAGS+HEALTH_TEMPLATE_TAGS},
                      QUARTERLY_IDX: {PC_tag: PC_QUARTERLY_TAGS+COMMON_QUARTERLY_TAGS,
                                      LIFE_tag: LIFE_QUARTERLY_TAGS+COMMON_QUARTERLY_TAGS,
                                      HEALTH_tag: HEALTH_QUARTERLY_TAGS+COMMON_QUARTERLY_TAGS}}

BUSINESS_TYPES = [PC_tag, LIFE_tag, HEALTH_tag]

TS_TEMPLATES = {PC_tag: 'PC_template', LIFE_tag: 'LIFE_template', HEALTH_tag: 'HEALTH_template'}
TS_TEMPLATES_Q = {PC_tag: 'PC_template_Q', LIFE_tag: 'LIFE_template_Q', HEALTH_tag: 'HEALTH_template_Q'}

DO_NOT_DISPLAY = 0x00
DISPLAY_TS_SECTION = 0x01
DISPLAY_TS_PROJ = 0x02
DISPLAY_MPL = 0x04
DISPLAY_DEBUG = 0x08

MODEL_PREV_YEAR = 1
MODEL_SGL_EXP = 2
MODEL_DBL_EXP = 3

ASSET_TYPE_STOCK = 0
ASSET_TYPE_BOND = 1
ASSET_TYPE_BA = 2
ASSET_TYPE_BA_ACQ = 3
ASSET_TYPE_BA_DISP = 4
ASSET_TYPE_STK_BOND_ACQ = 5
ASSET_TYPE_STK_BOND_DISP = 6

CUSIP_LIST = []
def get_etf_cusips():
    global CUSIP_LIST
    csv_data_dir = DATA_DIR + ETF_DIR
    files = os.listdir(csv_data_dir)
    for filename in files:
        if filename.lower().startswith('~'):
            continue
        if filename.lower().endswith('.csv'):
            full_filename = csv_data_dir + "\\" + filename
            rv_df = pd.read_csv(full_filename, header=0, index_col=0, dtype='unicode', thousands=",")
            rv_df['CUSIP8'] = rv_df['CUSIP'].str[:-1]
            CUSIP_LIST = CUSIP_LIST + rv_df.loc[ :, 'CUSIP8'].tolist()

    return

def do_all_calculations(template_wb, section_map, data_cube, position_dict, desired_periods):
    get_etf_cusips()
    ai_info = []
    bi_info = []
    ci_info = []
    di_info = []
    ei_info = []
    vi_info = []
    zi_info = []
    for tag, section in section_map.items():
        ai_info.append((tag, section.comp_dict_ai, section.fid_collection_dict))
        bi_info.append((tag, section.comp_dict_bi, section.fid_collection_dict))
        ci_info.append((tag, section.comp_dict_ci, section.fid_collection_dict))
        di_info.append((tag, section.comp_dict_di, section.fid_collection_dict))
        ei_info.append((tag, section.comp_dict_ei, section.fid_collection_dict))
        vi_info.append((tag, section.comp_dict_vi, section.fid_collection_dict))
        zi_info.append((tag, section.comp_dict_zi, section.fid_collection_dict))
    dependency_ordered_list = [vi_info, zi_info, ai_info, bi_info, ci_info,
                               di_info, ei_info]

    for calculation_level in dependency_ordered_list:
        for calc in calculation_level:
            data_cube = do_calculation(calc, template_wb, data_cube, section_map, position_dict, desired_periods)
    return data_cube


def do_calculation(calc, template_wb, data_cube, section_map, position_dict, desired_periods):
    df_dict = {}
    tag = calc[0]
    comp_dict = calc[1]
    giveup = False
    fid_collection_dict = calc[2]
    for key, fids in comp_dict.items():
        if desired_periods[0].find('Q') != -1:
            cell = key.replace('A', CALC_Q_COL)
        else:
            cell = key.replace('A', CALC_COL)
        formula = template_wb.get_formula(tag, cell)
        formula = formula.replace('$', '')
        func_idx = formula.find('(')
        args_idx = formula.rfind(')')
        func_key = formula[:func_idx]
        args_str = formula[func_idx + 1:args_idx]
        args = args_str.split(',')
        if func_key == AI_SET:
            slice1 = func_dict[func_key](args, data_cube)
        else:
            cell_info = args.pop(0)
            fid1 = lookup_fid(section_map, fid_collection_dict, cell_info)
            if fid1 is None:
                continue

            slice1 = data_cube[fid1]
            if func_key in POSITION_CALCS:
                slice1 = func_dict[func_key](args, position_dict, slice1, desired_periods)
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


def slice_set(args, cube):
    tmp_1 = cube.axes[1]
    tmp_2 = cube.axes[2]
    rv_slice = pd.DataFrame(float(args[0]), index=tmp_1, columns=tmp_2)
    return rv_slice


def two_year_ave(slice1):
    df = slice1
    df = df[df.columns[::-1]]  # reverse
    df = df.rolling(2, axis=1)
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

def bond_reflection(slice1, slice2):
    slice2 = slice2.loc[slice1.index]
    slice3 = slice2.div(slice1, axis=0)
    slice3 = slice3.iloc[:, 0:2].mean(axis=1)
    slice3 = slice1.iloc[:, 2:4].mul(slice3, axis=0)
    slice3 = pd.concat([slice2.iloc[:, 0:2], slice3], axis=1)
    slice3.fillna(0, inplace=True)
    return slice3


def bond_sum(args, position_dict, slice1, desired_periods):
    region = args.pop(0)
    slice2 = asset_sum_help(args, ASSET_TYPE_BOND, position_dict, desired_periods,
                           ['1', '2', '3', '4', '5', '6'], region)
    slice3 = bond_reflection(slice1, slice2)
    return slice3

def ig_sum(args, position_dict, slice1, desired_periods):
    region = args.pop(0)
    slice2 = asset_sum_help(args, ASSET_TYPE_BOND, position_dict, desired_periods, ['1', '2'], region)
    slice3 = bond_reflection(slice1, slice2)
    return slice3


def hy_sum(args, position_dict, slice1, desired_periods):
    region = args.pop(0)
    slice2 = asset_sum_help(args, ASSET_TYPE_BOND, position_dict, desired_periods, ['3', '4', '5', '6'], region)
    slice3 = bond_reflection(slice1, slice2)
    return slice3



def equity_etf_sum(args, position_dict, slice1, desired_periods):
    slice2 = asset_sum_help(args, ASSET_TYPE_STOCK, position_dict, desired_periods, ETF_STRINGS, CUSIP_LIST)
    slice3 = bond_reflection(slice1, slice2)
    return slice3

def equity_sum(args, position_dict, slice1, desired_periods):
    strings = []
    slice2 = asset_sum_help(args, ASSET_TYPE_STOCK, position_dict, desired_periods, strings, strings)
    return slice2

def ba_sum(args, position_dict, slice1, desired_periods):
    strings = []
    if args[0] == 'ACQ':
        asset_type = ASSET_TYPE_BA_ACQ
    elif args[0] == 'DISP':
        asset_type = ASSET_TYPE_BA_DISP
    else:
        asset_type = ASSET_TYPE_BA
    slice2 = asset_sum_help(args, asset_type, position_dict, desired_periods, strings, strings)
    return slice2

def liq_etf_sum(args, position_dict, slice1, desired_periods):
    what = args[0]
    if what == 'ACQ':
        asset_type = ASSET_TYPE_STK_BOND_ACQ
    elif what == 'DISP':
        asset_type = ASSET_TYPE_STK_BOND_DISP
    else:
        logger.error("undefined asset_type %s", args[0])
        return
    slice2 = asset_sum_help(args, asset_type, position_dict, desired_periods, "ETF", CUSIP_LIST)
    return slice2

def liq_sum(args, position_dict, slice1, desired_periods):
    what = args[0]
    region = args[1]
    if what == 'ACQ':
        asset_type = ASSET_TYPE_STK_BOND_ACQ
    elif what == 'DISP':
        asset_type = ASSET_TYPE_STK_BOND_DISP
    else:
        logger.error("undefined asset_type %s", args[0])
        return
    slice2 = asset_sum_help(args, asset_type, position_dict, desired_periods,
                            ['1','2','3','4','5','6'], region)
    return slice2

def liq_hy_sum(args, position_dict, slice1, desired_periods):
    what = args[0]
    region = args[1]
    if what == 'ACQ':
        asset_type = ASSET_TYPE_STK_BOND_ACQ
    elif what == 'DISP':
        asset_type = ASSET_TYPE_STK_BOND_DISP
    else:
        logger.error("undefined asset_type %s", args[0])
        return
    slice2 = asset_sum_help(args, asset_type, position_dict, desired_periods,
                            ['3','4','5','6'], region)
    return slice2

def liq_ig_sum(args, position_dict, slice1, desired_periods):
    what = args[0]
    region = args[1]
    if what == 'ACQ':
        asset_type = ASSET_TYPE_STK_BOND_ACQ
    elif what == 'DISP':
        asset_type = ASSET_TYPE_STK_BOND_DISP
    else:
        logger.error("undefined asset_type %s", args[0])
        return
    slice2 = asset_sum_help(args, asset_type, position_dict, desired_periods,
                            ['1','2'], region)
    return slice2

def asset_sum_help(args, asset_type, position_dict, desired_periods, strings1, strings2):
    locale.setlocale(locale.LC_NUMERIC, '')
    df_list = []
    price_sum = None
    asset_dict = None
    for period in desired_periods:
        if period in position_dict:
            if asset_type == ASSET_TYPE_BOND:
                asset_dict = position_dict[period][0]
            elif asset_type == ASSET_TYPE_STOCK:
                asset_dict = position_dict[period][1]
            elif asset_type == ASSET_TYPE_BA_ACQ:
                asset_dict = position_dict[period][0]
            elif asset_type == ASSET_TYPE_BA_DISP:
                asset_dict = position_dict[period][1]
            elif asset_type == ASSET_TYPE_STK_BOND_ACQ:
                asset_dict = position_dict[period][2]
            elif asset_type == ASSET_TYPE_STK_BOND_DISP:
                asset_dict = position_dict[period][3]
            elif asset_type == ASSET_TYPE_BA:
                if args[0] == 'ACQ':
                    asset_dict = position_dict[period][0]
                else:
                    asset_dict = position_dict[period][1]
            keys = list(asset_dict.keys())
            sum_list = []
            period_dict = {}
            for grp, df in asset_dict.items():
                tmp_df = df.copy(deep=True)
                if asset_type == ASSET_TYPE_STOCK:
                    price_sum = equity_sum_help(args, strings1, strings2, tmp_df)
                elif asset_type == ASSET_TYPE_BOND:
                    price_sum = bond_sum_help(args, strings1, strings2, tmp_df)
                elif asset_type == ASSET_TYPE_BA:
                    price_sum = ba_sum_help(args, strings1, strings2, tmp_df)
                elif asset_type == ASSET_TYPE_BA_ACQ:
                    price_sum = ba_sum_help(args, strings1, strings2, tmp_df)
                elif asset_type == ASSET_TYPE_BA_DISP:
                    price_sum = ba_sum_help(args, strings1, strings2, tmp_df)
                elif asset_type == ASSET_TYPE_STK_BOND_ACQ:
                    price_sum = liq_sum_help(args, strings1, strings2, tmp_df)
                elif asset_type == ASSET_TYPE_STK_BOND_DISP:
                    price_sum = liq_sum_help(args, strings1, strings2, tmp_df)
                sum_list.append(price_sum)
            period_dict[period] = sum_list
            df = pd.DataFrame.from_dict(period_dict)
            df.index = keys
            df_list.append(df)
        else:
             tmp_dict = {period: [0]}
             df = pd.DataFrame.from_dict(tmp_dict)
             df.index = ['Empty']
             df_list.append(df)

    df = pd.concat(df_list, axis=1)
    df.fillna(0, inplace=True)
    if df is not None:
        df = df.astype(long)
        if 'Empty' in df.index:
            df = df.drop('Empty')
    return df

# def ba_sum_help(args, string1, string2, tmp_df):
#     price_sum = 0.0
#     return price_sum

def bond_sum_help(args, naic_levels, region, tmp_df):
    if region == 'US':
        tmp_df = tmp_df.loc[tmp_df[ANNUAL_BONDS_FOREIGN_CODE_FID] == '...']
    elif region == 'EM':
        tmp_df = tmp_df.loc[tmp_df[ANNUAL_BONDS_FOREIGN_CODE_FID] == 'EM']
    elif region == 'OTHER':
        tmp_df = tmp_df.loc[tmp_df[ANNUAL_BONDS_FOREIGN_CODE_FID] != '...']
        tmp_df = tmp_df.loc[tmp_df[ANNUAL_BONDS_FOREIGN_CODE_FID] != 'EM']
    elif region == 'PUBLIC':
        tmp_df = tmp_df.loc[tmp_df[ANNUAL_BONDS_PRIVATE_PLACEMENT_FID] == '...']
        pass
    elif region == 'PRIVATE':
        tmp_df = tmp_df.loc[tmp_df[ANNUAL_BONDS_PRIVATE_PLACEMENT_FID] == 'X']
        pass
    # now slice by NAIC value
    tmp_df = tmp_df.loc[tmp_df[ANNUAL_BONDS_STMT_SEC_FID].isin(args)]
    tmp_df['naic_level'] = tmp_df[ANNUAL_BONDS_NAIC_DESIGNATION_FID].apply(lambda x: x[:1])
    tmp_df = tmp_df.loc[tmp_df['naic_level'].isin(naic_levels)]
    tmp_df[ANNUAL_BONDS_PRICE_FID] = tmp_df[ANNUAL_BONDS_PRICE_FID].apply(locale.atof)
    price_sum = tmp_df[ANNUAL_BONDS_PRICE_FID].sum() / 1000
    return price_sum

def equity_sum_help(args, strings, CUSIP_LIST, tmp_df):
    tmp_df = tmp_df.loc[tmp_df[ANNUAL_STOCKS_STMT_SEC_FID].isin(args)]
    if len(CUSIP_LIST) > 0:
        tmp_df = tmp_df[tmp_df['cusip'].apply(lambda x: any(y in x for y in CUSIP_LIST)) ]
    elif len(strings) > 0:
        tmp_df = tmp_df[
            (tmp_df[ANNUAL_STOCKS_ISSUER_NAME_FID].apply(lambda x: any(y in x for y in strings))) |
            (tmp_df[ANNUAL_STOCKS_ISSUER_DESC_FID].apply(lambda x: any(y in x for y in strings)))]
    if tmp_df.size == 0:
        price_sum = 0
    else:
        tmp_df[ANNUAL_STOCKS_PRICE_FID] = tmp_df[ANNUAL_STOCKS_PRICE_FID].apply(locale.atof)
        price_sum = tmp_df[ANNUAL_STOCKS_PRICE_FID].sum()/1000
    return price_sum

def ba_sum_help(args, strings, CUSIP_LIST, tmp_df):
    what = args[0]
    if what == 'ACQ':
        stmt_sec_fid = QUARTERLY_BA_ACQ_STMT_SEC_FID
    else:
        stmt_sec_fid = QUARTERLY_BA_DISP_STMT_SEC_FID

    tmp_df = tmp_df.loc[tmp_df[stmt_sec_fid].isin(args[1:])]
    if tmp_df.size == 0:
        price_sum = 0
    else:
        if what == 'ACQ':
            tmp_df[QUARTERLY_BA_ACQ_INITIAL_COST_FID] = tmp_df[QUARTERLY_BA_ACQ_INITIAL_COST_FID].apply(locale.atof)
            tmp_df[QUARTERLY_BA_ACQ_ADDITIONAL_COST_FID] = tmp_df[QUARTERLY_BA_ACQ_ADDITIONAL_COST_FID].apply(locale.atof)
            price_sum = tmp_df[QUARTERLY_BA_ACQ_INITIAL_COST_FID].sum()/1000 + \
                        tmp_df[QUARTERLY_BA_ACQ_ADDITIONAL_COST_FID].sum()/1000
        else:
            tmp_df[QUARTERLY_BA_DISP_CONSIDERATION_FID] = tmp_df[QUARTERLY_BA_DISP_CONSIDERATION_FID].apply(locale.atof)
            price_sum = tmp_df[QUARTERLY_BA_DISP_CONSIDERATION_FID].sum()/1000
    return price_sum

def liq_sum_help(args, strings1, strings2, tmp_df):
    what = args[0]
    sections = args[2:]
    if what == 'ACQ':
        stmt_sec_fid = QUARTERLY_STK_BOND_ACQ_STMT_SEC_FID
        naic_level_fid = QUARTERLY_STK_BOND_ACQ_NAIC_FID
        foreign_fid = QUARTERLY_STK_BOND_ACQ_FOREIGN_FID
        asset_desc_fid = QUARTERLY_STK_BOND_ACQ_ASSET_DESC_FID
    else:
        stmt_sec_fid = QUARTERLY_STK_BOND_DISP_STMT_SEC_FID
        naic_level_fid = QUARTERLY_STK_BOND_DISP_NAIC_FID
        foreign_fid = QUARTERLY_STK_BOND_DISP_FOREIGN_FID
        asset_desc_fid = QUARTERLY_STK_BOND_DISP_ASSET_DESC_FID

    if len(sections) > 0:
        if sections[0] < '84':
            naic_levels = strings1
            region = strings2
            if region == 'US':
                tmp_df = tmp_df.loc[tmp_df[foreign_fid] == '...']
            elif region == 'EM':
                tmp_df = tmp_df.loc[tmp_df[foreign_fid] == 'EM']
            elif region == 'OTHER':
                tmp_df = tmp_df.loc[tmp_df[foreign_fid] != '...']
                tmp_df = tmp_df.loc[tmp_df[foreign_fid] != 'EM']

            tmp_df['naic_level'] = tmp_df[naic_level_fid].apply(lambda x: x[:1])
            tmp_df = tmp_df.loc[tmp_df['naic_level'].isin(naic_levels)]
        elif strings1 == "ETF":
            if len(CUSIP_LIST) > 0:
                tmp_df = tmp_df[tmp_df['cusip'].apply(lambda x: any(y in x for y in CUSIP_LIST))]
            elif len(ETF_STRINGS) > 0:
                tmp_df = tmp_df[ tmp_df[asset_desc_fid].apply(lambda x: any(y in x for y in ETF_STRINGS))]

    tmp_df = tmp_df.loc[tmp_df[stmt_sec_fid].isin(sections)]
    if tmp_df.size == 0:
        price_sum = 0
    else:
        if what == 'ACQ':
            tmp_df[QUARTERLY_STK_BOND_ACQ_INITIAL_COST_FID] = tmp_df[QUARTERLY_STK_BOND_ACQ_INITIAL_COST_FID].apply(locale.atof)
            tmp_df[QUARTERLY_STK_BOND_ACQ_ADDITIONAL_COST_FID] = tmp_df[QUARTERLY_STK_BOND_ACQ_ADDITIONAL_COST_FID].apply(locale.atof)
            price_sum = tmp_df[QUARTERLY_STK_BOND_ACQ_INITIAL_COST_FID].sum()/1000 + \
                        tmp_df[QUARTERLY_STK_BOND_ACQ_ADDITIONAL_COST_FID].sum()/1000
        else:
            tmp_df[QUARTERLY_STK_BOND_DISP_CONSIDERATION_FID] = tmp_df[QUARTERLY_STK_BOND_DISP_CONSIDERATION_FID].apply(locale.atof)
            price_sum = tmp_df[QUARTERLY_STK_BOND_DISP_CONSIDERATION_FID].sum()/1000
    return price_sum

AI_SUM = "=AI_SUM"
AI_DIFF = "=AI_DIFF"
AI_MULT = "=AI_MULT"
AI_DIV = "=AI_DIV"

AI_HY = "=AI_HY"
AI_IG = "=AI_IG"
AI_BOND_SUM = "=AI_BOND_SUM"

AI_EQ = "=AI_EQ"
AI_EQ_ETF = "=AI_EQ_ETF"
AI_LIQ_ETF = "=AI_LIQ_ETF"

AI_BA = "=AI_BA"
AI_LIQ = "=AI_LIQ"
AI_HY_LIQ = "=AI_HY_LIQ"
AI_IG_LIQ = "=AI_IG_LIQ"

AI_SET_DF = "=AI_SET_DF"
AI_SET = "=AI_SET"
AI_2YR_AVE = "=AI_2YR_AVE"

POSITION_CALCS = [AI_LIQ_ETF, AI_EQ_ETF, AI_EQ, AI_HY, AI_IG, AI_BOND_SUM, AI_BA, AI_LIQ, AI_HY_LIQ, AI_IG_LIQ]

func_dict = {
    AI_SUM: slice_sum,
    AI_DIFF: slice_diff,
    AI_MULT: slice_mult,
    AI_DIV: slice_div,
    AI_BOND_SUM: bond_sum,
    AI_IG: ig_sum,
    AI_HY: hy_sum,
    AI_BA: ba_sum,
    AI_LIQ: liq_sum,
    AI_LIQ_ETF: liq_etf_sum,
    AI_HY_LIQ: liq_hy_sum,
    AI_IG_LIQ: liq_ig_sum,
    AI_EQ: equity_sum,
    AI_EQ_ETF: equity_etf_sum,
    AI_SET: slice_set,
    AI_SET_DF: slice_set_df,
    AI_2YR_AVE: two_year_ave
     }


class Section(object):
    def __init__(self, tag, fid_collection_dict, comp_dict_ai, comp_dict_bi,
                 comp_dict_ci, comp_dict_di, comp_dict_ei,
                 comp_dict_vi, comp_dict_zi):
        self.tag = tag
        self.fid_collection_dict = fid_collection_dict
        self.comp_dict_ai = comp_dict_ai
        self.comp_dict_bi = comp_dict_bi
        self.comp_dict_ci = comp_dict_ci
        self.comp_dict_di = comp_dict_di
        self.comp_dict_ei = comp_dict_ei
        self.comp_dict_vi = comp_dict_vi
        self.comp_dict_zi = comp_dict_zi
    pass
