# partial filenames for acquiring data
Assets_files    = "Assets_ALL"
CashFlow_files  = "CashFlow_ALL"
E07_files       = "E07_ALL"
E10_files       = "E10_ALL"
LiabSurp_files  = "LiabSurp_ALL"
SI01_files      = "SI01_ALL"
SI05_07_files   = "SI05_07_ALL"
SI08_09_files   = "SI08_09_ALL"
SoI_files       = "SoI_ALL"

# template for creating tearsheet
template_locations = ""
template_filename = "TearSheet_Template.xlsx"
TEMPLATE_SHEETS = ["E07", "SI01"]
#TEMPLATE_SHEETS = ["Assets", "CashFlow", "E07", "E10", "LiabSurp", "SI01", "SI05_07", "SI08_09", "SoI"]
TEMPLATE_ROW_LABELS_E07 = 'B7:B54'
SI01_idx = 0
E07_idx = 1

target_filename = "TearSheet_Output.xlsx"
target_sheet = "TS"
pc_tag = "PC"
life_tag = "LIFE"
health_tag = "HEALTH"

needed_sheets = ["E07", "SI01"]
