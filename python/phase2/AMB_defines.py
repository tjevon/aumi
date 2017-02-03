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

PC_tag      = "PC"
LIFE_tag    = "LIFE"
HEALTH_tag  = "HEALTH"

CALC_COL = 'E'

COMMON_TEMPLATE_TAGS = [Assets_tag, CashFlow_tag, SI01_tag, SI05_07_tag, E07_tag]
PC_TEMPLATE_TAGS = [SoI_tag, IRIS2_tag]
LIFE_TEMPLATE_TAGS = [SoO_tag, IRIS1_tag]
HEALTH_TEMPLATE_TAGS = [SoR_tag]

PERCENT_FORMATS = [IRIS1_tag, IRIS2_tag]

TEMPLATE_FILENAME = "TearSheet_Template.xlsx"

TARGET_FILENAME = "TearSheet_Output.xlsx"
TARGET_SHEET = "TS"

def sum(list, fid_collection_dict, cube_dict):
    fid1 = fid_collection_dict[list[0]][0]
    slice = cube_dict[fid1].copy()
    for row in list[1:]:
        fid2 = fid_collection_dict[row][0]
        slice2 = cube_dict[fid2]
        slice += slice2
    return slice

func_dict = {"=AI_SUM": sum }