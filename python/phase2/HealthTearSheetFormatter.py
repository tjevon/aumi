from __future__ import print_function
import os

import xlwings as xw

import PcBusinessType as pc
import LifeBusinessType as life
import HealthBusinessType as health

import logging
from AMB_defines import *
from TearSheetFormatter import *
logger = logging.getLogger('twolane')

class HealthTearSheetFormatter(TearSheetFormatter):
    def __init__(self, template_wb, target_wb, pandas_xl_writer):
        super(HealthTearSheetFormatter, self).__init__(template_wb, target_wb, pandas_xl_writer)
        pass

    def create_tearsheets(self, companies, mpl, line_no):
        for co in companies:
            for tag in HEALTH_TEMPLATE_TAGS:
                line_no = self.format_section(co, mpl, tag, line_no)
            super(HealthTearSheetFormatter, self).create_tearsheets(co, mpl, line_no)


    def get_business_type(self,mpl):
        return mpl.business_types[HEALTH_tag]