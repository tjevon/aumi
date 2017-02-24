from __future__ import print_function

import xlwings as xw

from TearSheetFormatter import *

logger = logging.getLogger('twolane')


class TearSheetGenerator:
    """Class to manage generation of tear sheets in excel"""
    def __init__(self, data_dir, template_obj):
        logger.debug("Enter")
        self.pandas_xl_write = None
        self.target_wb = None
        self.template_obj = template_obj
        self.get_target_workbook(data_dir)

        self.ts_formatter = TearSheetFormatter(self.template_obj, self.target_wb, self.pandas_xl_write)

        logger.debug("Leave")

    def get_target_workbook(self, data_dir):
        output_data_dir = data_dir + "\\output"
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
        for i in self.target_wb.sheets:
            i.clear_contents()

    def get_target_worksheet(self, sheet_name):
        target_sheet = self.target_wb.sheets(sheet_name)
        return target_sheet

    def add_xlsheets(self, sheet_list, y_or_q):
        available_sheets = []
        for i in self.target_wb.sheets:
            logger.debug(i.name)
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        if y_or_q == QUARTERLY_IDX:
            sheet_list = [x + '_Q' for x in sheet_list]

        for entry in sheet_list:
            entry_sheet = entry
            if entry_sheet not in available_sheets:
                self.target_wb.sheets.add(entry_sheet)

    def build_tearsheets(self, company_dict, mpl, y_or_q):
        line_no = 4
        if PC_tag in company_dict:
            self.add_xlsheets(company_dict[PC_tag], y_or_q)
            self.ts_formatter.create_tearsheets(company_dict, PC_tag, mpl, line_no, y_or_q)
        if LIFE_tag in company_dict:
            self.add_xlsheets(company_dict[LIFE_tag], y_or_q)
            self.ts_formatter.create_tearsheets(company_dict, LIFE_tag, mpl, line_no, y_or_q)
        if HEALTH_tag in company_dict:
            self.add_xlsheets(company_dict[HEALTH_tag], y_or_q)
            self.ts_formatter.create_tearsheets(company_dict, HEALTH_tag, mpl, line_no, y_or_q)
        pass
