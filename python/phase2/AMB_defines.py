import pandas as pd

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

CALC_COL = 'E'

COMMON_TEMPLATE_TAGS = [Assets_tag, CashFlow_tag, SI01_tag, SI05_07_tag, E07_tag, CR_tag]
PC_TEMPLATE_TAGS = [SoI_tag, IRIS1_tag]
LIFE_TEMPLATE_TAGS = [SoO_tag, IRIS2_tag, Liab2_tag]
HEALTH_TEMPLATE_TAGS = [SoR_tag, Liab3_tag]

PERCENT_FORMATS = [IRIS1_tag, IRIS2_tag]

TEMPLATE_FILENAME = "TearSheet_Template.xlsx"

TARGET_FILENAME = "TearSheet_Output.xlsx"
TARGET_SHEET = "TS"

def do_all_calculations(template_wb, section_map, cube):
    ai_info = []
    bi_info = []
    ci_info = []
    for tag, section in section_map.iteritems():
        ai_info.append((tag, section.comp_dict_ai, section.fid_collection_dict))
        bi_info.append((tag, section.comp_dict_bi, section.fid_collection_dict))
        ci_info.append((tag, section.comp_dict_ci, section.fid_collection_dict))
    dependency_ordered_list = [ai_info, bi_info, ci_info]

    for calculation_level in dependency_ordered_list:
        for calc in calculation_level:
            cube = do_calculation(calc, template_wb, cube)
    return cube

def do_calculation(calc, template_wb, cube):
    df_dict = {}
    tag = calc[0]
    comp_dict = calc[1]
    fid_collection_dict = calc[2]
    for key, fids in comp_dict.iteritems():
        cell = key.replace('A', CALC_COL)
        formula = template_wb.get_formula(tag, cell)
        func_idx = formula.find('(')
        args_idx = formula.rfind(')')
        func_key = formula[:func_idx]
        args_str = formula[func_idx + 1:args_idx]
        args = args_str.split(",")
        fid1 = fid_collection_dict[args[0]][0]
        slice1 = cube[fid1]
        for row in args[1:]:
            fid2 = fid_collection_dict[row][0]
            slice2 = cube[fid2]
            slice1 = func_dict[func_key](slice1, slice2)
        df_dict[fids[0]] = slice1
    tmp_cube = pd.Panel(df_dict)
    cube_list = [cube, tmp_cube]
    cube = pd.concat(cube_list, axis=0)
    return cube


def slice_sum(slice1, slice2):
    rv_slice = slice1 + slice2
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

func_dict = {
    "=AI_SUM": slice_sum,
    "=AI_DIFF": slice_diff,
    "=AI_MULT": slice_mult,
    "=AI_DIV": slice_div
     }
