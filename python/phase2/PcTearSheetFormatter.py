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

class PcTearSheetFormatter(TearSheetFormatter):
    def __init__(self, template_obj, target_wb, pandas_xl_writer):
        super(PcTearSheetFormatter, self).__init__(template_obj, target_wb, pandas_xl_writer)
        pass

    # concrete methods
    def get_business_type(self,mpl):
        return mpl.business_types[pc_tag]
