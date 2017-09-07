from __future__ import print_function

import xlwings as xw


from TearSheetFormatter import *

logger = logging.getLogger('twolane')


class TearSheetGenerator:
    """Class to manage generation of tear sheets in excel"""
    def __init__(self, template_obj):
        logger.debug("Enter")
        self.pandas_xl_write = None
        self.target_wb = None
        self.template_obj = template_obj
        self.get_target_workbook()

        self.ts_formatter = TearSheetFormatter(self.template_obj, self.target_wb, self.pandas_xl_write)

        logger.debug("Leave")

    def get_target_workbook(self):
        output_data_dir = XL_DIR + OUTPUT_DIR
        workbook_file = output_data_dir + "\\" + TARGET_FILENAME
        try:
            self.target_wb = xw.Book(workbook_file)
            pass
        except:
            pass
            self.target_wb = xw.books.add()
            self.target_wb.save(workbook_file)
        try:
            # self.pandas_xl_write =  pd.ExcelWriter(workbook_file,engine='xlsxwriter')
            pass
        except:
            logger.error("Pandas ExcelWriter error: %s", workbook_file)
            pass
        self.target_wb.app.calculation = 'manual'

    def get_target_worksheet(self, sheet_name):
        target_sheet = self.target_wb.sheets(sheet_name)
        return target_sheet

    def add_xlsheets(self, sheet_list, bt_tag, y_or_q, debug=False):
        available_sheets = []
        for i in self.target_wb.sheets:
            logger.debug(i.name)
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        sheet_list.sort()

        for entry in sheet_list:
            logger.error("Tag %s Name %s", bt_tag, entry)
            entry_sheet = entry[:30] if len(entry) > 30 else entry
            entry_sheet = entry_sheet.translate(None, "".join(BAD_CHAR))

            if entry_sheet not in available_sheets:
                if debug == True:
                    self.target_wb.sheets.add(entry_sheet)
                    continue

                if y_or_q == QUARTERLY_IDX:
                    before = self.target_wb.sheets('Sheet1')
                    self.template_obj.get_template_sheet(TS_TEMPLATES[bt_tag]).api.Copy(Before=before.api)
                    self.target_wb.sheets(TS_TEMPLATES[bt_tag]).api.Name = entry_sheet
                    self.target_wb.sheets(entry_sheet).range('A3').value = entry
                elif y_or_q == MPL_IDX:
                    before = self.target_wb.sheets('Sheet1')
                    self.template_obj.get_template_sheet(entry_sheet).api.Copy(Before=before.api)
                    self.target_wb.sheets(entry_sheet).api.Name = entry_sheet
                elif y_or_q == IF_IDX:
                    before = self.target_wb.sheets('Sheet1')
                    self.template_obj.get_template_sheet('Sector_Info').api.Copy(Before=before.api)
                    self.target_wb.sheets('Sector_Info').api.Name = entry_sheet
                elif bt_tag != '':
                    before = self.target_wb.sheets('Sheet1')
                    self.template_obj.get_template_sheet(TS_TEMPLATES[bt_tag]).api.Copy(Before=before.api)
                    self.target_wb.sheets(TS_TEMPLATES[bt_tag]).api.Name = entry_sheet
                    self.target_wb.sheets(entry_sheet).range('A3').value = entry
        return
    def read_stone_harbor(self, filename):
        df = pd.read_csv(filename, header=None, dtype='unicode')
        cusip_list = df.iloc[:,1].tolist()
        cusip_list = [x for x in cusip_list if str(x) != 'nan']
        cusip_list = [x[:-1] for x in cusip_list]
        return cusip_list

    def build_cusip_ownership_sheet(self, asset_mgr, mpl):
        bt_grp_dict = {}
        csv_data_dir = DATA_DIR + asset_mgr
        files = os.listdir(csv_data_dir)
        for file in files:
            csv_filename = csv_data_dir + "\\" + file
            if asset_mgr == 'StoneHarbor':
                cusip_list = self.read_stone_harbor(csv_filename)
                bt_grp_dict = mpl.get_groups_owning_cusips(cusip_list)
        logger.debug("Leave")
        return

    def build_tearsheets(self, company_dict, mpl, y_or_q, debug=False):
        if PC_tag in company_dict and company_dict[PC_tag] is not None:
            self.add_xlsheets(company_dict[PC_tag].values(), PC_tag, y_or_q, debug)
            self.ts_formatter.create_tearsheets(company_dict[PC_tag], PC_tag, mpl, y_or_q, debug)
        if LIFE_tag in company_dict and company_dict[LIFE_tag] is not None:
            self.add_xlsheets(company_dict[LIFE_tag].values(), LIFE_tag, y_or_q, debug)
            self.ts_formatter.create_tearsheets(company_dict[LIFE_tag], LIFE_tag, mpl, y_or_q, debug)
        if HEALTH_tag in company_dict and company_dict[HEALTH_tag] is not None:
            self.add_xlsheets(company_dict[HEALTH_tag].values(), HEALTH_tag, y_or_q, debug)
            self.ts_formatter.create_tearsheets(company_dict[HEALTH_tag], HEALTH_tag, mpl, y_or_q, debug)
        pass

