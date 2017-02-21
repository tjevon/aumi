from __future__ import print_function

import logging
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

    def build_column_labels(self, years):
        column_heading = []
        num_years = len(years)
        num_cols = (num_years - 1) + num_years + 1
        col_idx = 0
        year_idx = 0
        num_years = len(years)
        years = sorted(years, reverse=True)
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

    def build_row_labels(self, tag, bus_type_tag):
        row_labels = self.template_obj.get_row_labels(tag, bus_type_tag)
        return row_labels

    def build_df_for_display(self, co, tag, bus_type):
        template_fids_with_spaces = self.template_obj.get_display_fid_list(tag, bus_type)
        just_fids = filter(lambda a: a is not None, template_fids_with_spaces)
        just_fids = filter(lambda a: a != 'XXX', just_fids)

        df = bus_type.get_df_including_pcts(co, just_fids)
        split_idx = len(bus_type.years)
        l1 = df.columns[0:split_idx]
        l2 = df.columns[split_idx:]
        cols = [val for pair in zip(l1, l2) for val in pair]
        df = df.reindex(columns=cols)

        blank_row = ["" for x in range(len(df.columns))]
        tup = tuple(blank_row)
        blank_row_df = pd.DataFrame([tup], columns=df.columns)
        rv_df = pd.DataFrame(columns=df.columns)
        mat_line = 0
        for row in template_fids_with_spaces:
            if row is None or row == 'XXX' or row == '' or row == ' ':
                rv_df = pd.concat([rv_df, blank_row_df])
                continue
            rv_df = pd.concat([rv_df, df[mat_line:mat_line+1]])
            mat_line += 1
            pass
        return rv_df

    def format_section(self, co, mpl, tag, line_no):
        logger.info("Enter")
        bus_type = self.get_business_type(mpl)

        cell = ('B', line_no)
        column_labels = self.build_column_labels(bus_type.years)
        self.copy_labels_to_xlsheet(column_labels, co, cell)

        cell = ('A', line_no)
        row_labels = self.build_row_labels(tag, bus_type.get_bt_tag())
        self.copy_labels_to_xlsheet(row_labels, co, cell)

        cell = ('B', line_no + 1)
        df = self.build_df_for_display(co, tag, bus_type)
        self.copy_df_to_xlsheet(df, co, cell, tag)

        logger.info("Leave")
        return df.shape[0] + line_no + 3

    def copy_labels_to_xlsheet(self, values, co, cell_info):
        cell = cell_info[0] + str(cell_info[1])
        target_sheet = co
#        target_sheet = co + "_" + TARGET_SHEET
        target_sheet = self.target_wb.sheets(target_sheet)
        target_sheet.range(cell).value = values
        target_sheet.range('A1:A500').autofit()
        return

    def copy_df_to_xlsheet(self, df, co, cell_info, tag):
        target_sheet = co
#        target_sheet = co + "_" + TARGET_SHEET
        target_sheet = self.target_wb.sheets(target_sheet)
        df = df.replace([np.inf,-np.inf],np.nan)
        cell = cell_info[0] + str(cell_info[1])
        target_sheet.range(cell).options(dropna=False, index=False, header=False).value = df
        height = len(df.index)
        top = cell_info[1]
        bottom = cell_info[1] + height
        #        col = chr(ord(cell_info[0])+1)
        col = cell_info[0]
        for i in range(len(df.columns)):
            left_cell = col + str(top)
            right_cell = col + str(bottom)
            xl_range = left_cell + ":" + right_cell
            if i % 2 == 0:
                if tag in PERCENT_FORMATS:
                    target_sheet.range(xl_range).number_format = '0.0'
                else:
                    target_sheet.range(xl_range).number_format = '$0'
            else:
                target_sheet.range(xl_range).number_format = '0.00%'
            target_sheet.range(xl_range).column_width = 10
            col = chr(ord(col)+1)
        return

    def get_business_type(self,mpl):
        ### pure virtual ###
        pass
