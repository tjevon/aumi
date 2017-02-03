from __future__ import print_function
import os

import xlwings as xw

import PcBusinessType as pc
import LifeBusinessType as life
import HealthBusinessType as health

import logging
import pandas as pd
import numpy as np
from AMB_defines import *

logger = logging.getLogger('twolane')

class TearSheetFormatter(object):
    def __init__(self, template_obj, target_wb, pandas_xl_writer):
        self.template_obj = template_obj
        self.target_wb = target_wb
        self.pandas_xl_writer = pandas_xl_writer
        pass

    def create_tearsheets(self, co, mpl, line_no):
        for tag in COMMON_TEMPLATE_TAGS:
            line_no = self.format_section(co, mpl, tag, line_no)
        return line_no

    def format_section(self, co, mpl, tag, line_no):
        logger.info("Enter")
        my_bt = self.get_business_type(mpl)

        my_cell = ('B', line_no )
        num_years = len(my_bt.years)
        num_cols = (num_years - 1) + num_years + 1
        column_heading = self.get_column_headings(my_bt.years, num_cols)
        self.copy_labels(column_heading, co, my_cell)

        my_cell = ('A', line_no + 1 )
        row_labels = self.template_obj.get_row_labels(tag, my_bt.get_bt_tag())
        self.copy_labels(row_labels, co, my_cell)


        template_fids_with_spaces = self.template_obj.get_display_fid_list(tag, my_bt)
        just_fids = filter(lambda a: a != None, template_fids_with_spaces)
        just_fids = filter(lambda a: a != 'XXX', just_fids)

        my_df = my_bt.get_df_including_pcts(co, tag)
        my_df = my_df.reindex(just_fids)

        my_cell = ('B', line_no + 1 )
        l_df = self.setup_df_data(template_fids_with_spaces, my_df)
        self.copy_df_data(l_df, co, my_cell, tag)
        return len(l_df.index) + line_no + 5

    def setup_df_data(self, fids, df):
        l_blank_row = ["" for x in range(len(df.columns))]
        l_tup = tuple(l_blank_row)
        l_blank_row_df = pd.DataFrame([l_tup], columns=df.columns)
        l_df = pd.DataFrame(columns=df.columns)
        mat_line = 0
        df_add_blank_row_list = [l_df, l_blank_row_df]
        for row in fids:
            if row == None or row == 'XXX' or row == '' or row == ' ':
                l_df = pd.concat([l_df, l_blank_row_df])
                continue
            l_df = pd.concat([l_df, df[mat_line:mat_line+1]])
            mat_line += 1
            pass
        return l_df

    def get_column_headings(self, years, num_cols):
        column_heading = []
        col_idx = 0
        year_idx = 0
        num_years = len(years)
        years = sorted(years,reverse=False)
        while col_idx < num_cols-1 and year_idx < num_years:
            if (col_idx % 2) == 0:
                column_heading.append(years[year_idx])
                year_idx += 1
            else:
                tmp_str = "% Chg"
                column_heading.append(tmp_str)
            col_idx += 1

        tmp_str = "%d Yr %% Chg" % (num_years - 1)
        column_heading.append(tmp_str)
        return column_heading

    def get_business_type(self,mpl):
        ### pure virtual ###
        pass

    def copy_labels(self, values, co, cell_info):
        my_cell = cell_info[0] + str(cell_info[1])
        my_target_sheet = co + "_" + TARGET_SHEET
        my_target_sheet = self.target_wb.sheets(my_target_sheet)
        my_target_sheet.range(my_cell).value = values
        my_target_sheet.range('A1:A500').autofit()
        return

    def copy_df_data(self, df, co, cell_info, tag):
        my_target_sheet = co + "_" + TARGET_SHEET
        my_target_sheet = self.target_wb.sheets(my_target_sheet)
        df = df.replace([np.inf,-np.inf],np.nan)
        my_cell = cell_info[0] + str(cell_info[1])
        my_target_sheet.range(my_cell).options(dropna=False, index=False, header=False).value = df
        height = len(df.index)
        top = cell_info[1]
        bottom = cell_info[1] + height
        #        col = chr(ord(cell_info[0])+1)
        col = cell_info[0]
        for i in range(len(df.columns)):
            l_cell = col + str(top)
            r_cell = col + str(bottom)
            xl_range = l_cell + ":" + r_cell
            if i % 2 == 0:
                if tag in PERCENT_FORMATS:
                    my_target_sheet.range(xl_range).number_format = '0.0'
                else:
                    my_target_sheet.range(xl_range).number_format = '$0'
            else:
                my_target_sheet.range(xl_range).number_format = '0.00%'
            my_target_sheet.range(xl_range).column_width = 10
            col = chr(ord(col)+1)
        return

