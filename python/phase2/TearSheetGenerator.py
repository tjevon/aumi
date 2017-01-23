from __future__ import print_function
import os

import xlwings as xw

import PcBusinessType as pc
import LifeBusinessType as life
import HealthBusinessType as health

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')

class BaseTearSheetFormatter(object):
    target_col_labels_SI01 = 'B3'
    target_row_labels_SI01 = 'A4'
    target_col_labels_E07 = 'B69'
    target_row_labels_E07 = 'A70'
    target_data = 'B70'

    template_row_labels_SI01 = 'A4:A67'
    template_row_labels_E07 = 'B7:B54'

    def __init__(self, template_wb, target_wb):
        self.template_wb = template_wb
        self.target_wb = target_wb
        pass

    def create_tearsheets(self, companies, mpl):
        for co in companies:
            self.format_SI01_section(co, mpl)
            self.format_E07_section(co, mpl)
        pass

    def format_SI01_section(self, co, mpl):
        logger.info("Enter")
        my_bt = self.get_business_type(mpl)

        my_template_sheet = self.template_wb.sheets(template_sheets[SI01_idx])
        my_values = my_template_sheet.range(self.template_row_labels_SI01).options(ndim=2).value
        self.copy_labels(my_values, co, self.target_row_labels_SI01)

        num_years = len(my_bt.years)
        num_cols = (num_years - 1) + num_years + 2
        column_heading = self.get_column_headings(my_bt.years, num_cols)
        self.copy_labels(column_heading, co, self.target_col_labels_SI01)
        pass

    def create_fid_collection(self, fids, num_years):
        fid_collection = [["" for x in range(num_years)] for y in range(len(fids))]
        for fid, row in zip(fids, range(len(fids))):
            if fid[0] == None:
                continue
            for col in num_years:
                if col == 0:
                    fid_collection[row][col] = fid[0]
                else:
                    fid_collection[row][col] = fid[0] + "." + str(col)
        return fid_collection

    def create_fid_collection_old(self, fids, num_years, num_cols):
        fid_collection = []
        fid_collection = [["" for x in range(num_cols)] for y in range(len(fids))]
        for fid, row in zip(fids, range(len(fids))):
            if fid[0] == None:
                continue
            col = 0
            x = 0
            while col < num_cols - 1:
                if col == 0:
                    fid_one = fid[0]
                    x = 0
                else:
                    fid_one = fid[0] + "." + str(x)
                fid_collection[row][col] = fid_one
                col += 1
                if col < num_cols - 2:
                    fid_two = fid[0] + "." + str(x) + "." + str(x+1)
                    x += 1
                    fid_collection[row][col] = fid_two
                col += 1

            fid_two = fid[0] + "." + str(x) + "." + str(0)
            fid_collection[row][col-1] = fid_two
        return fid_collection

    def calc_percent_change(self, fid_collection, df):
        for row in fid_collection:
            if row[0] == '':
                continue
            x = 0
            y = 1
            z = 2
            while z < len(row):
                df[row[y]] = df[row[x]]/df[row[z]] - 1.0
                x += 2
                y += 2
                z += 2
                pass
            df[row[y]] = df[row[0]]/df[row[x]] - 1.0

        df.fillna(0.0,inplace=True)
        pass

    def format_E07_section(self, co, mpl):
        logger.info("Enter")
        my_bt = self.get_business_type(mpl)

        my_template_sheet = self.template_wb.sheets(TEMPLATE_SHEETS[E07_idx])
        my_values = my_template_sheet.range(self.template_row_labels_E07).options(ndim=2).value
        self.copy_labels(my_values, co, self.target_row_labels_E07)

        num_years = len(my_bt.years)
        num_cols = (num_years - 1) + num_years + 1
        column_heading = self.get_column_headings(my_bt.years, num_cols)
        self.copy_labels(column_heading, co, self.target_col_labels_E07)

        my_fids = my_template_sheet.range(self.get_E07_column()).options(ndim=2).value
#        fid_collection = self.create_fid_collection_old(my_fids, num_years, num_cols)
        fid_collection = self.create_fid_collection(my_fids, num_years)
        df_columns = my_bt.E07_raw_df.columns.values.tolist()
        self.calc_percent_change(fid_collection, my_bt.E07_raw_df)

        row_label = co
        num_rows = len(my_fids)
        the_mat = [[0.0 for x in range(num_cols)] for y in range(num_rows)]
        mat_row_idx = 0
        for fid_row in fid_collection:
            for fid, mat_col_idx in zip( reversed(fid_row), range(0,len(fid_row))):
                if fid == '':
                    the_mat[mat_row_idx][mat_col_idx] = ""
                    continue
                if mat_col_idx == 0:
                    the_mat[mat_row_idx][num_cols-1] = my_bt.E07_raw_df.loc[row_label,fid]
                else:
                    the_mat[mat_row_idx][mat_col_idx-1] = my_bt.E07_raw_df.loc[row_label, fid]
            mat_row_idx += 1

        self.copy_data(the_mat, co, self.target_data )
        pass
        return the_mat

#        self.create_matrix(fids, co, mpl)

    def get_column_headings(self, years, num_cols):
        column_heading = []
        col_idx = 0
        year_idx = 0
        num_years = len(years)
        count = 0
        while col_idx < num_cols and year_idx < num_years:
            if col_idx < 2 or (col_idx % 2) == 1:
                column_heading.append(years[year_idx])
                year_idx += 1
            else:
                tmp_str = "%s%% Chg" % (years[year_idx - 1])
                column_heading.append(tmp_str)
            col_idx += 1

        tmp_str = "%s%% Chg" % (years[year_idx - 1])
        column_heading.append(tmp_str)
        tmp_str = "%d Yr %% Chg" % (num_years - 1)
        column_heading.append(tmp_str)
        tmp_str = "Annualized % Chg"
        column_heading.append(tmp_str)
        return column_heading

    def get_business_type(self,mpl):
        pass


    def get_E07_column(self):
        return 'Z1'

    def copy_col_labels(self, values, co, target_range):
        pass

    def copy_labels(self, values, co, target_range):
        my_target_sheet = co + "_" + target_sheet
        my_target_sheet = self.target_wb.sheets(my_target_sheet)
        my_target_sheet.range(target_range).value = values
        my_target_sheet.autofit()

    def copy_data(self, values, co, target_range):
        my_target_sheet = co + "_" + target_sheet
        my_target_sheet = self.target_wb.sheets(my_target_sheet)
        my_target_sheet.range(target_range).value = values
        my_target_sheet.autofit()

class PcTearSheetFormatter(BaseTearSheetFormatter):
    def __init__(self, template_wb, target_wb):
        super(PcTearSheetFormatter, self).__init__(template_wb, target_wb)
        pass

    def create_pc_tearsheet(self, fids, co, mpl):
        my_bt = self.get_business_type(mpl)
        mat_col_idx = 0
        mat_line_idx = 0
#        num_cols = (num_years - 1) + num_years + 2
#        num_rows = len(fin_sub_tots[0]) + len(fin_sub_tots[1]) + len(fin_sub_tots[2]) + 3
#        fin_mat = [[0.0 for x in range(num_cols)] for y in range(num_rows)]

#        num_years = len(my_bt.years)
#        year_idx = 0
#        while year_idx > 0
#            for i in range(len(my_bt.years),0):
#                if
#                    my_bt.
        pass

    # concrete methods
    def get_business_type(self,mpl):
        return mpl.business_types[pc_tag]
        pass
    def get_E07_column(self):
        return 'C7:C54'

class LifeTearSheetFormatter(BaseTearSheetFormatter):
    def __init__(self, template_wb, target_wb):
        super(LifeTearSheetFormatter, self).__init__(template_wb, target_wb)
        pass

class HealthTearSheetFormatter(BaseTearSheetFormatter):
    def __init__(self, template_wb, target_wb):
        super(HealthTearSheetFormatter, self).__init__(template_wb, target_wb)
        pass

class TearSheetGenerator:
    """Class to manage generation of tear sheets in excel"""
    def __init__(self, data_dir):
        logger.debug("Enter")
        self.get_template_workbook(data_dir)
        self.get_target_workbook(data_dir)

        self.pc_formatter = PcTearSheetFormatter(self.template_wb, self.target_wb)
        self.life_formatter = LifeTearSheetFormatter(self.template_wb, self.target_wb)
        self.health_formatter = HealthTearSheetFormatter(self.template_wb, self.target_wb)

        logger.debug("Leave")

    def get_template_workbook(self, data_dir):
        template_data_dir = data_dir + "\\templates"
        workbook_file = template_data_dir + "\\" + template_filename
        self.template_wb = xw.Book(workbook_file)
        available_sheets = []
        for i in self.template_wb.sheets:
            logger.debug(i.name)
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        for sheet_name in template_sheets:
            if sheet_name not in available_sheets:
                logger.fatal("Template %s does not contain %s", workbook_file, sheet_name)
                exit()

    def get_target_workbook(self, data_dir):
        output_data_dir = data_dir + "\\output"
        workbook_file = output_data_dir + "\\" + target_filename
        try:
            self.target_wb = xw.Book(workbook_file)
        except:
            self.target_wb = xw.books.add()
            self.target_wb.save(workbook_file)

    def add_xlsheets(self, companies):
        available_sheets = []
        for i in self.target_wb.sheets:
            logger.debug(i.name)
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        for co in companies:
            co_sheet = co + "_" + target_sheet
            if co_sheet not in available_sheets:
                self.target_wb.sheets.add(co_sheet)

    def build_pc_tearsheets(self, companies, mpl):
        logger.debug("Enter: num companies = %d", len(companies))
        self.add_xlsheets(companies)
        self.pc_formatter.create_tearsheets(companies, mpl)
        logger.debug("Leave")
        pass

    def build_tearsheets(self, symbol_list, business):
        pass
