import os

DATA_DIR = os.getenv('DATAPATH', './')
XL_DIR = os.getenv('XLPATH', './')

YEARLY_DIR = "V1_Annual_data"
QTRLY_DIR = "V1_Quarterly_data"

COMPANY_INFO_DIR = "V1_company_info"
COMPANY_MAP_DIR = "V1_company_mapping"

POSITIONS_DIR = "V1_inv_unaf_af"
ETF_DIR = "ETF_cantor"

TEMPLATE_DIR = "templates"
OUTPUT_DIR = "output"

TEMPLATE_FILENAME = "TearSheet_Template.xlsx"
TARGET_FILENAME = "TearSheet_Output.xlsx"
