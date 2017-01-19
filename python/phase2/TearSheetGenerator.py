from __future__ import print_function
import os

import xlwings as xw

import PcBusinessType as pc
import LifeBusinessType as life
import HealthBusinessType as health

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')

class BaseTearSheetFormatter(object):
    SI01_cell = "A3"
    E07_cell = "A67"
    def __init__(self, template_wb, target_wb):
        self.template_wb = template_wb
        self.target_wb = target_wb
        pass

    def format_SI01(self,companies, mpl):
        my_template_sheet = self.template_wb.sheets(template_sheets[SI01_idx])
        self.copy_row_labels(my_template_sheet,'A4:A67', companies, 'A4')
        pass

    def format_E07(self,companies, mpl):
        my_template_sheet = self.template_wb.sheets(template_sheets[E07_idx])
        self.copy_row_labels(my_template_sheet,'B7:B54', companies, 'A69')
        my_fids = my_template_sheet.range(self.get_E07_column()).options(ndim=2).value
        pass

    def get_E07_column(self):
        return 'Z1'

    def copy_row_labels(self, template_sheet, template_range, companies, target_range):
        for co in companies:
            my_target_sheet = co + "_" + target_sheet
            my_target_sheet = self.target_wb.sheets(my_target_sheet)
            my_values = template_sheet.range(template_range).options(ndim=2).value
            my_target_sheet.range(target_range).value = my_values
            my_target_sheet.autofit()

class PcTearSheetFormatter(BaseTearSheetFormatter):
    def __init__(self, template_wb, target_wb):
        super(PcTearSheetFormatter, self).__init__(template_wb, target_wb)
        pass

    def get_E07_column(self):
        return 'C7:C54'

    def print_tearsheets(self, companies, mpl):
        super(PcTearSheetFormatter, self).format_SI01(companies, mpl)
        super(PcTearSheetFormatter, self).format_E07(companies, mpl)
        pass

class LifeTearSheetFormatter(BaseTearSheetFormatter):
    def __init__(self, template_wb, target_wb):
        super(LifeTearSheetFormatter, self).__init__(template_wb, target_wb)
        pass

class HealthTearSheetFormatter(BaseTearSheetFormatter):
    def __init__(self, template_wb, target_wb):
        super(HealthTearSheetFormatter, self).__init__(template_wb, target_wb)
        pass

class TearSheetGenerator:
    """Class to manage generation of tear sheets in excel"""
    def __init__(self, data_dir):
        logger.debug("Enter")
        self.get_template_workbook(data_dir)
        self.get_target_workbook(data_dir)

        self.pc_formatter = PcTearSheetFormatter(self.template_wb, self.target_wb)
        self.life_formatter = LifeTearSheetFormatter(self.template_wb, self.target_wb)
        self.health_formatter = HealthTearSheetFormatter(self.template_wb, self.target_wb)

        logger.debug("Leave")

    def get_template_workbook(self, data_dir):
        template_data_dir = data_dir + "\\templates"
        workbook_file = template_data_dir + "\\" + template_filename
        self.template_wb = xw.Book(workbook_file)
        available_sheets = []
        for i in self.template_wb.sheets:
            logger.debug(i.name)
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        for sheet_name in template_sheets:
            if sheet_name not in available_sheets:
                logger.fatal("Template %s does not contain %s", workbook_file, sheet_name)
                exit()

    def get_target_workbook(self, data_dir):
        output_data_dir = data_dir + "\\output"
        workbook_file = output_data_dir + "\\" + target_filename
        try:
            self.target_wb = xw.Book(workbook_file)
        except:
            self.target_wb = xw.books.add()
            self.target_wb.save(workbook_file)

    def create_sheets(self, companies):
        available_sheets = []
        for i in self.target_wb.sheets:
            logger.debug(i.name)
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        for co in companies:
            co_sheet = co + "_" + target_sheet
            if co_sheet not in available_sheets:
                self.target_wb.sheets.add(co_sheet)

    def build_pc_tearsheets(self, companies, mpl):
        logger.debug("Enter: num companies = %d", len(companies))
        self.create_sheets(companies)
        self.pc_formatter.print_tearsheets(companies, mpl)
        logger.debug("Leave")
        pass

    def build_tearsheets(self, symbol_list, business):
        pass
