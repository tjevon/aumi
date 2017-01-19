# partial filenames for acquiring data
sch_d      = "SchD_ALL"
assets     = "Assets_ALL"
cash_flow  = "CashFlow_ALL"
sis        = "SIS_ALL"
soi        = "SOI_ALL"
sch_ba     = "SchBA_ALL"

# template for creating tearsheet
template_locations = ""
template_filename = "TearSheet_Template.xlsx"
template_sheets = ["SI01", "E07"]
SI01_idx = 0
E07_idx = 1

target_filename = "TearSheet_Output.xlsx"
target_sheet = "TS"
pc_tag = "PC"
life_tag = "LIFE"
health_tag = "HEALTH"

needed_sheets = [sch_d, assets, cash_flow, sis]
