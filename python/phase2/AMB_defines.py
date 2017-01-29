# partial filenames for acquiring data
Assets_files    = "Assets_ALL"
CashFlow_files  = "CashFlow_ALL"
E07_files       = "E07_ALL"
E10_files       = "E10_ALL"
LiabSurp_files  = "LiabSurp_ALL"
SI05_07_files   = "SI05_07_ALL"
SI08_09_files   = "SI08_09_ALL"
SoI_files       = "SoI_ALL"

# template for creating tearsheet
SI01_tag = "SI01"
E07_tag = "E07"
Assets_tag = "Assets"
SoI_tag = "SoI"
CashFlow_tag = "CashFlow"

SI01_files      = SI01_tag + "_ALL"

COMMON_TEMPLATE_SHEETS = [E07_tag, SI01_tag, Assets_tag, CashFlow_tag]
PC_TEMPLATE_SHEETS = [SoI_tag]
#TEMPLATE_ROW_LABELS_E07 = 'B7:B54'
E07_idx = 0
SI01_idx = 1
#Assets_idx = 2

template_locations = ""
template_filename = "TearSheet_Template.xlsx"

target_filename = "TearSheet_Output.xlsx"
target_sheet = "TS"

pc_tag = "PC"
life_tag = "LIFE"
health_tag = "HEALTH"
