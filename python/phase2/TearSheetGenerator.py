from __future__ import print_function
import os

import xlwings as xw

import PcBusinessType as pc
import LifeBusinessType as life
import HealthBusinessType as health

import logging
from AMB_defines import *
from PcTearSheetFormatter import *
from LifeTearSheetFormatter import *
from HealthTearSheetFormatter import *

logger = logging.getLogger('twolane')


class TearSheetGenerator:
    """Class to manage generation of tear sheets in excel"""
    def __init__(self, data_dir, template_obj):
        logger.debug("Enter")
        self.pandas_xl_write = None
        self.target_wb = None
        self.template_obj = template_obj
        self.get_target_workbook(data_dir)

        self.pc_formatter = PcTearSheetFormatter(self.template_obj, self.target_wb, self.pandas_xl_write)
        self.life_formatter = LifeTearSheetFormatter(self.template_obj, self.target_wb, self.pandas_xl_write)
        self.health_formatter = HealthTearSheetFormatter(self.template_obj, self.target_wb, self.pandas_xl_write)

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
#            self.pandas_xl_write =  pd.ExcelWriter(workbook_file,engine='xlsxwriter')
            pass
        except:
            logger.error("Pandas ExcelWriter error: %s", workbook_file)
            pass
        available_sheets = []
        for i in self.target_wb.sheets:
            i.clear_contents()


    def add_xlsheets(self, companies):
        available_sheets = []
        for i in self.target_wb.sheets:
            logger.debug(i.name)
            #            i.clear_contents()
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        for co in companies:
            co_sheet = co + "_" + TARGET_SHEET
            if co_sheet not in available_sheets:
                self.target_wb.sheets.add(co_sheet)

    def build_pc_tearsheets(self, companies, mpl, line_no):
        logger.debug("Enter: num companies = %d", len(companies))
        self.add_xlsheets(companies)
        self.pc_formatter.create_tearsheets(companies, mpl, line_no)
        logger.debug("Leave")
        pass

    def build_life_tearsheets(self, companies, mpl, line_no):
        logger.debug("Enter: num companies = %d", len(companies))
        self.add_xlsheets(companies)
        self.life_formatter.create_tearsheets(companies, mpl, line_no)
        logger.debug("Leave")
        pass

    def build_health_tearsheets(self, companies, mpl, line_no):
        logger.debug("Enter: num companies = %d", len(companies))
        self.add_xlsheets(companies)
        self.health_formatter.create_tearsheets(companies, mpl, line_no)
        logger.debug("Leave")
        pass

    def build_tearsheets(self, company_dict, mpl):
        line_no = 4
        if company_dict[PC_tag] != None:
            self.build_pc_tearsheets(company_dict[PC_tag], mpl, line_no)
        if company_dict[LIFE_tag] != None:
            self.build_life_tearsheets(company_dict[LIFE_tag], mpl, line_no)
        if company_dict[HEALTH_tag] != None:
            self.build_health_tearsheets(company_dict[HEALTH_tag], mpl, line_no)
        pass
