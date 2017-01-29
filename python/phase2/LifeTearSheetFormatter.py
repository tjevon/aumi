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

class LifeTearSheetFormatter(TearSheetFormatter):
    def __init__(self, template_wb, target_wb):
        super(LifeTearSheetFormatter, self).__init__(template_wb, target_wb)
        pass

    def get_business_type(self,mpl):
        return mpl.business_types[life_tag]