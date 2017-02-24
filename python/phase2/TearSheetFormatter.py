from __future__ import print_function

import logging
import numpy as np
import itertools
import string
from AMB_defines import *

logger = logging.getLogger('twolane')

class TearSheetFormatter(object):
    def __init__(self, template_obj, target_wb, pandas_xl_writer):
        self.template_obj = template_obj
        self.target_wb = target_wb
        self.pandas_xl_writer = pandas_xl_writer
        pass

    def create_tearsheets(self, company_dict, bt_tag, mpl, line_no, y_or_q):
        initial_line = line_no
        for co in company_dict[bt_tag]:
            line_no = initial_line
            for tag in mpl.business_types[bt_tag][y_or_q].bt_tag_list:
                line_no = self.format_section(co, bt_tag, mpl, tag, line_no, y_or_q)
            for tag in mpl.business_types[bt_tag][y_or_q].common_tag_list:
                line_no = self.format_section(co, bt_tag, mpl, tag, line_no, y_or_q)
        return line_no

    def build_column_labels(self, periods):
        column_heading = []
        num_periods = len(periods)
        num_cols = (num_periods - 1) + num_periods + 1
        col_idx = 0
        period_idx = 0
        while col_idx < num_cols-1 and period_idx < num_periods:
            if (col_idx % 2) == 0:
                column_heading.append(periods[period_idx])
                period_idx += 1
            else:
                tmp_str = "% Chg"
                column_heading.append(tmp_str)
            col_idx += 1

        tmp_str = "%d Yr %% Chg" % (num_periods - 1)
        column_heading.append(tmp_str)
        return column_heading

    def build_row_labels(self, tag, bus_type_tag):
        row_labels = self.template_obj.get_row_labels(tag, bus_type_tag)
        return row_labels

    def build_df_for_display(self, co, tag, bus_type, y_or_q):
        template_fids_with_spaces = self.template_obj.get_display_fid_list(tag, bus_type, y_or_q)
        just_fids = filter(lambda a: a is not None, template_fids_with_spaces)
        just_fids = filter(lambda a: a != 'XXX', just_fids)

        df = bus_type.get_df_including_pcts(co, just_fids)
        split_idx = len(bus_type.periods)
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

    def format_section(self, co, bt_tag, mpl, tag, line_no, y_or_q):
        logger.info("Enter")
        bus_type = self.get_business_type(mpl, bt_tag, y_or_q)
        page = co
        if y_or_q == QUARTERLY_IDX:
            page = co + '_Q'

        cell = ('B', line_no)
        column_labels = self.build_column_labels(bus_type.periods)
        self.copy_labels_to_xlsheet(column_labels, page, cell)

        cell = ('A', line_no)
        row_labels = self.build_row_labels(tag, bus_type.get_bt_tag())
        self.copy_labels_to_xlsheet(row_labels, page, cell)

        cell = ('B', line_no + 1)
        df = self.build_df_for_display(co, tag, bus_type, y_or_q)
        self.copy_df_to_xlsheet(df, page, cell, tag)

        logger.info("Leave")
        return df.shape[0] + line_no + 3

    def copy_labels_to_xlsheet(self, values, co, cell_info):
        cell = cell_info[0] + str(cell_info[1])
        target_sheet = co
        target_sheet = self.target_wb.sheets(target_sheet)
        target_sheet.range(cell).value = values
        target_sheet.range('A1:A500').autofit()
        return

    def copy_df_to_xlsheet(self, df, co, cell_info, tag):
        target_sheet = co
        target_sheet = self.target_wb.sheets(target_sheet)
        df = df.replace([np.inf,-np.inf],np.nan)
        cell = cell_info[0] + str(cell_info[1])
        target_sheet.range(cell).options(dropna=False, index=False, header=False).value = df
        height = len(df.index)
        top = cell_info[1]
        bottom = cell_info[1] + height
        #        col = chr(ord(cell_info[0])+1)
        col = cell_info[0]
        strings = [''.join(letters) for length in xrange(1, 3) for letters in
                   itertools.product(string.ascii_uppercase, repeat=length)]
        col_idx = strings.index(col)
        strings = strings[col_idx:]
        for i, c in zip(range(len(df.columns)), strings):
            left_cell = c + str(top)
            right_cell = c + str(bottom)
            xl_range = left_cell + ":" + right_cell
            if i % 2 == 0:
                if tag in PERCENT_FORMATS:
                    target_sheet.range(xl_range).number_format = '0.0'
                else:
                    target_sheet.range(xl_range).number_format = '$0'
            else:
                target_sheet.range(xl_range).number_format = '0.00%'
            target_sheet.range(xl_range).column_width = 10
        return

    def get_business_type(self,mpl, bt_tag, y_or_q):
        return mpl.business_types[bt_tag][y_or_q]
