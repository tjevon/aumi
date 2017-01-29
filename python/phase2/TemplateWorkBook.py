from __future__ import print_function
import os

import xlwings as xw

import PcBusinessType as pc
import LifeBusinessType as life
import HealthBusinessType as health

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')

class TemplateWorkBook(object):
    def __init__(self, data_dir):
        self.xl_wb = self.get_template_workbook(data_dir)
#        self.E07_template_sheet = self.xl_wb.sheets(COMMON_TEMPLATE_SHEETS[E07_idx])
#        self.SI01_template_sheet = self.xl_wb.sheets(COMMON_TEMPLATE_SHEETS[SI01_idx])
#        self.Assets_template_sheet = self.xl_wb.sheets(COMMON_TEMPLATE_SHEETS[Assets_idx])
#        self.SoI_template_sheet = self.xl_wb.sheets(PC_TEMPLATE_SHEETS[0])

    def get_template_workbook(self, data_dir):
        template_data_dir = data_dir + "\\templates"
        workbook_file = template_data_dir + "\\" + template_filename
        template_wb = xw.Book(workbook_file)
        available_sheets = []
        for i in template_wb.sheets:
            logger.debug(i.name)
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        for sheet_name in COMMON_TEMPLATE_SHEETS:
            if sheet_name not in available_sheets:
                logger.fatal("Template %s does not contain %s", workbook_file, sheet_name)
                exit()
        return template_wb

    def get_template_sheet(self, tag):
        the_template_sheet = self.xl_wb.sheets(tag)
        return the_template_sheet

