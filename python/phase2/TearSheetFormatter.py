from __future__ import print_function

import numpy as np
from AMB_defines import *

logger = logging.getLogger('twolane')

ADDRESS_FID =   'CO00030'
CITY_FID =      'CO00033'
STATE_FID =     'CO00034'
ZIP_FID =       'CO00035'
PHONE_FID =     'CO00020'
WEBSITE_FID =   'CO00117'
FAX_FID =       'CO00022'

CIO_FID =       'CO00329'
CFO_FID =       'CO00327'
CEO_FID =       'CO00326'

CIO_EMAIL_FID = 'CO00330'
CFO_EMAIL_FID = 'CO00331'
CEO_EMAIL_FID = 'CO00332'

BUSINESS_FOCUS_FID = 'CO00179'

DEBUG_START_LINE = 85

class TearSheetFormatter(object):
    def __init__(self, template_obj, target_wb, pandas_xl_writer):
        self.template_obj = template_obj
        self.target_wb = target_wb
        self.pandas_xl_writer = pandas_xl_writer
        self.debug_row = DEBUG_START_LINE
        pass

    @staticmethod
    def get_period_type(mpl, bt_tag, y_or_q):
        return mpl.business_types[bt_tag].period_types[y_or_q]

    @staticmethod
    def build_column_labels(periods):
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

    def format_header_section(self, co_num, co_name, bt_tag, mpl):
        target_sheet = co_name[:30] if len(co_name) > 30 else co_name
        target_sheet = target_sheet.translate(None, "".join(BAD_CHAR))
        target_sheet = self.target_wb.sheets(target_sheet)

        comp_df = mpl.business_types[bt_tag].company_info_df.loc[co_num, :]
        address = comp_df[ADDRESS_FID] + ' ' + comp_df[CITY_FID] + ', ' + comp_df[STATE_FID] + ' ' + comp_df[ZIP_FID]

        target_sheet.range('B4').value = address
        target_sheet.range('B5').value = comp_df[PHONE_FID]
        target_sheet.range('B6').value = comp_df[WEBSITE_FID]
        target_sheet.range('B7').value = comp_df[FAX_FID]

        target_sheet.range('F4').value = comp_df[CEO_FID]
        target_sheet.range('F5').value = comp_df[CFO_FID]
        target_sheet.range('F6').value = comp_df[CIO_FID]
        target_sheet.range('F7').value = comp_df[BUSINESS_FOCUS_FID]

        target_sheet.range('J4').value = comp_df[CEO_EMAIL_FID]
        target_sheet.range('J5').value = comp_df[CFO_EMAIL_FID]
        target_sheet.range('J6').value = comp_df[CIO_EMAIL_FID]
        return


    def create_tearsheets(self, sheet_list, bt_tag, mpl, y_or_q, debug=False):
        for co_num, co_name in sheet_list.items():
            for tag in mpl.business_types[bt_tag].period_types[y_or_q].complete_tag_list:
                self.format_section(co_num, co_name, bt_tag, mpl, tag, y_or_q, debug)
#            if y_or_q == YEARLY_IDX:
                self.format_projections_section(co_num, co_name, bt_tag, Liquid_Assets_tag, mpl)
                self.format_projections_section(co_num, co_name, bt_tag, E07_tag, mpl)
                self.format_header_section(co_num, co_name, bt_tag, mpl)
            self.debug_row = DEBUG_START_LINE
        return

    def build_row_labels(self, tag, bus_type_tag, display_type, debug=False):
        label_info = self.template_obj.get_row_labels(tag, bus_type_tag, display_type, self.debug_row, debug)
        self.debug_row += len(label_info[0]) + 1

        return label_info

    def build_df_for_display(self, co, tag, period_type, y_or_q, debug=False):
        template_fids_with_spaces = self.template_obj.get_projection_info(tag, period_type.get_bt_tag(),
                                                                          y_or_q, DISPLAY_TS_SECTION, debug)[0]
        just_fids = filter(lambda a: a is not None, template_fids_with_spaces)
        just_fids = filter(lambda a: a != 'XXX', just_fids)

        df = period_type.get_df_including_pcts(co, just_fids)
        split_idx = len(period_type.desired_periods)
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

    def build_df_for_quarterly_display_2(self, co, tag, period_type, y_or_q, b_idx, sz, debug=False):

        template_fids_with_spaces = self.template_obj.get_projection_info(tag, period_type.get_bt_tag(),
                                                                          y_or_q, DISPLAY_TS_SECTION, debug)[0]
        just_fids = filter(lambda a: a is not None, template_fids_with_spaces)
        just_fids = filter(lambda a: a != 'XXX', just_fids)
        df = period_type.get_df_including_pcts_2(co, just_fids)


        return rv_df

    def build_df_for_quarterly_display(self, co, tag, period_type, y_or_q, b_idx, sz, debug=False):
        template_fids_with_spaces = self.template_obj.get_projection_info(tag, period_type.get_bt_tag(),
                                                                          y_or_q, DISPLAY_TS_SECTION, debug)[0]
        just_fids = filter(lambda a: a is not None, template_fids_with_spaces)
        just_fids = filter(lambda a: a != 'XXX', just_fids)

        df = period_type.get_df_including_pcts(co, just_fids)
        split_idx = len(period_type.desired_periods)

        l1 = df.columns[b_idx:b_idx+sz]
        l2 = df.columns[split_idx+b_idx:split_idx+b_idx+sz]
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

    def format_section(self, co_num, co_name, bt_tag, mpl, tag, y_or_q, debug=False):
        if y_or_q == YEARLY_IDX:
            self.format_annual_section(co_num, co_name, bt_tag, mpl, tag, y_or_q, debug)
        else:
            self.format_quarterly_section(co_num, co_name, bt_tag, mpl, tag, y_or_q, debug)

    def format_quarterly_section(self, co_num, co_name, bt_tag, mpl, tag, y_or_q, debug=False):
        logger.info("Enter")
        annual_obj = self.get_period_type(mpl, bt_tag, YEARLY_IDX)
        quarterly_obj = self.get_period_type(mpl, bt_tag, QUARTERLY_IDX)
        page = co_name

        desired_periods = ["YTD Change"] + annual_obj.desired_periods[:3]

        label_info = self.build_row_labels(tag, quarterly_obj.get_bt_tag(), DISPLAY_TS_SECTION, debug)
        column_labels = self.build_column_labels(desired_periods)

        if label_info[3] is None:
            return
        row_labels = label_info[0]
        row_label_col = label_info[1]
        data_column = label_info[2]
        row = int(label_info[3])

        cell = (data_column, row)
        self.copy_labels_to_xlsheet(column_labels, page, cell)

        cell = (row_label_col, row)
        self.copy_labels_to_xlsheet(row_labels, page, cell)

        df = self.build_df_for_quarterly_display(co_num, tag, quarterly_obj, y_or_q, 0,1, debug)
        cell = (data_column, row+1)
        self.copy_df_to_xlsheet(df, page, cell)

        df = self.build_df_for_quarterly_display(co_num, tag, annual_obj, 0, 1,3, debug)
        data_column = chr(ord(data_column)+2)
        cell = (data_column, row+1)
        self.copy_df_to_xlsheet(df, page, cell)

        logger.info("Leave")
        return

    def format_annual_section(self, co_num, co_name, bt_tag, mpl, tag, y_or_q, debug=False):
        logger.info("Enter")
        period_type = self.get_period_type(mpl, bt_tag, y_or_q)
        page = co_name

        label_info = self.build_row_labels(tag, period_type.get_bt_tag(), DISPLAY_TS_SECTION, debug)
        column_labels = self.build_column_labels(period_type.desired_periods)

        if label_info[3] is None:
            return
        row_labels = label_info[0]
        row_label_col = label_info[1]
        data_column = label_info[2]
        row = int(label_info[3])

        cell = (data_column, row)
        self.copy_labels_to_xlsheet(column_labels, page, cell)

        cell = (row_label_col, row)
        self.copy_labels_to_xlsheet(row_labels, page, cell)

        df = self.build_df_for_display(co_num, tag, period_type, y_or_q, debug)

        cell = (data_column, row+1)
        self.copy_df_to_xlsheet(df, page, cell)

        logger.info("Leave")
        return

    def format_projections_section(self, co_num, co_name, bt_tag, tag, mpl):
        page = co_name
        label_info = self.build_row_labels(tag, bt_tag, DISPLAY_TS_PROJ)

        row_labels = label_info[0]
        row_label_col = label_info[1]
        if row_label_col is None:
            return
        data_column = label_info[2]
        row = int(label_info[3])

        cell = (row_label_col, row)
        self.copy_labels_to_xlsheet(row_labels, page, cell)

        proj_info = self.template_obj.get_projection_info(tag, bt_tag, YEARLY_IDX, DISPLAY_TS_PROJ)
        fids_with_spaces = proj_info[0]
        just_fids = filter(lambda a: a is not None, fids_with_spaces)
        just_fids = filter(lambda a: a != 'XXX', just_fids)

        qtrly_proj_cube = mpl.qtrly_proj_dict[bt_tag]
        full_qtrly_df = qtrly_proj_cube.major_xs(co_num).transpose()
        full_yrly_df = mpl.yrly_proj_dict[bt_tag]
        full_yrly_df = full_yrly_df.loc[co_num, :]

        yrly_proj_df = full_yrly_df.loc[just_fids]
        qtrly_proj_df = full_qtrly_df.loc[just_fids, :]
        qtrly_proj_df = qtrly_proj_df.reindex(just_fids)

        data_df = mpl.business_types[bt_tag].period_types[YEARLY_IDX].get_df_including_pcts(co_num, just_fids)
        data_df = data_df.iloc[:, 0:1]

        sec_df = pd.concat([data_df, qtrly_proj_df, yrly_proj_df], axis=1)
        self.copy_df_to_xlsheet(sec_df, co_name, (data_column, row+1))
        return

    def copy_labels_to_xlsheet(self, values, co, cell_info):
        cell = cell_info[0] + str(cell_info[1])
        target_sheet = co[:30] if len(co) > 30 else co
        target_sheet = target_sheet.translate(None, "".join(BAD_CHAR))
        target_sheet = self.target_wb.sheets(target_sheet)
        target_sheet.range(cell).value = values
        return

    def copy_df_to_xlsheet(self, df, co, cell_info):
        target_sheet = co[:30] if len(co) > 30 else co
        target_sheet = target_sheet.translate(None, "".join(BAD_CHAR))
        target_sheet = self.target_wb.sheets(target_sheet)
        df = df.replace([np.inf, -np.inf], np.nan)
        cell = cell_info[0] + str(cell_info[1])
        target_sheet.range(cell).options(dropna=False, index=False, header=False).value = df
        return
