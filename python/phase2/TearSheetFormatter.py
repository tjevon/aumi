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
    # TODO: put template stuff in template
    template_row_labels = {
        SI01_tag:'A4:A67',
        E07_tag:'B7:B54',
        Assets_tag: 'A2:A58',
        CashFlow_tag: 'A6:A57'
    }
    target_col_labels = {
        SI01_tag:'B3',
        E07_tag:'B69',
        Assets_tag: 'B119',
        CashFlow_tag: 'B180'
    }
    target_row_labels = {
        SI01_tag:'A4',
        E07_tag:'A70',
        Assets_tag: 'A120',
        CashFlow_tag: 'A181'
    }
    target_data_labels = {
        SI01_tag:'B5',
        E07_tag:'B70',
        Assets_tag: 'B120',
        CashFlow_tag: 'B181'
    }

    pandas_target_data_labels = {
        SI01_tag:('B',5),
        E07_tag:('B',70),
        Assets_tag: ('B',120),
        CashFlow_tag: ('B',181)
    }

    def __init__(self, template_obj, target_wb, pandas_xl_writer):
        self.template_obj = template_obj
        self.target_wb = target_wb
        self.pandas_xl_writer = pandas_xl_writer
        pass

    def create_tearsheets(self, companies, mpl):
        for co in companies:
            self.format_section(co, mpl, SI01_tag)
            self.format_section(co, mpl, E07_tag)
            self.format_section(co, mpl, Assets_tag)
            self.format_section(co, mpl, CashFlow_tag)
        pass


    def format_section(self, co, mpl, tag):
        logger.info("Enter")
        my_bt = self.get_business_type(mpl)

        my_template_sheet = self.template_obj.get_template_sheet(tag)
        row_headings = my_template_sheet.range(self.template_row_labels[tag]).options(ndim=2).value
        self.copy_labels(row_headings, co, self.target_row_labels[tag])

        num_years = len(my_bt.years)
        num_cols = (num_years - 1) + num_years + 2
        column_heading = self.get_column_headings(my_bt.years, num_cols)
        self.copy_labels(column_heading, co, self.target_col_labels[tag])

        template_fids = my_template_sheet.range(my_bt.get_template_column(tag)).options(transpose=True).value
        just_fids = filter(lambda a: a != None or a != 'XXX', template_fids)
        my_df = my_bt.get_df(co, tag)

        my_df = my_df.reindex(just_fids)
        l_df = self.setup_df_data(template_fids, my_df)
        self.copy_df_data(l_df, co, self.pandas_target_data_labels[tag])
        pass

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

#    def get_row_labels(self, tag):
#        pass

    def get_column_headings(self, years, num_cols):
        column_heading = []
        col_idx = 0
        year_idx = 0
        num_years = len(years)
        while col_idx < num_cols-2 and year_idx < num_years:
            if (col_idx % 2) == 0:
                column_heading.append(years[year_idx])
                year_idx += 1
            else:
                tmp_str = "% Chg"
                column_heading.append(tmp_str)
            col_idx += 1

        tmp_str = "%d Yr %% Chg" % (num_years - 1)
        column_heading.append(tmp_str)
        tmp_str = "Annualized % Chg"
        column_heading.append(tmp_str)
        return column_heading

    def get_business_type(self,mpl):
        ### pure virtual ###
        pass

    def copy_labels(self, values, co, target_range):
        my_target_sheet = co + "_" + target_sheet
        my_target_sheet = self.target_wb.sheets(my_target_sheet)
        my_target_sheet.range(target_range).value = values
        my_target_sheet.autofit()

    def copy_df_data(self, df, co, tag):
        my_target_sheet = co + "_" + target_sheet
        my_target_sheet = self.target_wb.sheets(my_target_sheet)
        df = df.replace([np.inf,-np.inf],np.nan)
        cell = tag[0] + str(tag[1])
        my_target_sheet.range(cell).options(dropna=False, index=False, header=False).value = df
        height = len(df.index)
        top = tag[1]
        bottom = tag[1] + height
        #        col = chr(ord(tag[0])+1)
        col = tag[0]
        for i in range(len(df.columns)):
            l_cell = col + str(top)
            r_cell = col + str(bottom)
            xl_range = l_cell + ":" + r_cell
            if i % 2 == 0:
                my_target_sheet.range(xl_range).number_format = '$0'
            else:
                my_target_sheet.range(xl_range).number_format = '0.00%'
            my_target_sheet.range(xl_range).column_width = 10
            col = chr(ord(col)+1)

