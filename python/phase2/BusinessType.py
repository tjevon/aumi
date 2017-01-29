from __future__ import print_function
import pandas as pd

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')

class BusinessType(object):
    """Base class for PC, Life, Health"""
    def __init__(self):
        self.Assets_raw_df = None
        self.CashFlow_raw_df     = None
        self.E07_raw_df = None
        self.E10_raw_df = None
        self.SI01_raw_df    = None
        self.SI05_07_raw_df  = None

        self.Assets_cube = None
        self.CashFlow_cube = None
        self.E07_cube    = None
        self.E10_cube    = None
        self.SI01_cube    = None
        self.SI05_07_cube    = None

        self.data_cubes = {}

        self.years = []
        self.companies = set()
        pass

    def load_df(self, csv_filename):
        logger.info("Enter: %s", csv_filename)
        the_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode')
        self.get_years(the_df)

        # should read df with field # for column lable and group # for row label
        the_df = pd.read_csv(csv_filename, header=3, index_col=0, dtype='unicode')
        the_df = the_df.drop(['AMB#'])
        the_df = the_df.drop('Unnamed: 1', 1)
        the_df = the_df.drop('Unnamed: 2', 1)
        for i in the_df.index:
            if (i.find("Unnamed") != -1) or (i.find(".") != -1) or (i.find("AMB#") != -1):
                continue
            if pd.notnull(i):
                self.companies.add(i)
        for i in the_df.columns:
            if i.find('Calc') != -1:
                the_df[i].replace(regex=True,inplace=True,to_replace=r'%',value=r'')

        # try placing this above the for and eliminating if/continue
        the_df = the_df.astype(float)
        logger.info("Leave")
        return the_df

    def process_csvs( self, data_dir, file_names ):
        logger.info("Enter")

        return_file_names = []
        for file_name in file_names:
            csv_filename = data_dir + "\\" + file_name
            the_df = self.load_df(csv_filename)

            if file_name.find(Assets_files) != -1:
                self.Assets_raw_df = the_df
            elif file_name.find(CashFlow_files) != -1:
                self.CashFlow_raw_df = the_df
            elif file_name.find(E07_files) != -1:
                self.E07_raw_df = the_df
            elif file_name.find(E10_files) != -1:
                self.E10_raw_df = the_df
            elif file_name.find(SI01_files) != -1:
                self.SI01_raw_df = the_df
            elif file_name.find(SI05_07_files) != -1:
                self.SI05_07_raw_df = the_df
            else:
                return_file_names.append(file_name)
        logger.info("Leave")
        return return_file_names

    def construct_data_cubes(self, template_wb, sheet_list):
        for sheet in sheet_list:
            cube_dict = {}
            my_template_sheet = template_wb.xl_wb.sheets(sheet)
            my_fids = my_template_sheet.range(self.get_template_column(sheet)).options(ndim=2).value
            fid_collection = self.create_fid_collection(my_fids, self.years)
            the_df = None
            df_return = self.get_approp_df(sheet)
            if df_return[0] == False:
                logger.error("No df available in base class for sheet: %s", sheet)
                df_return = self.get_derived_df(sheet)
                if df_return[0] == False:
                    logger.error("No df available in derived class for sheet: %s", sheet)
                    continue
            the_df = df_return[1]
            for fids in fid_collection:
                if fids[0] == '' or fids[0] == 'XXX':
                    continue
                cube_slice_df = the_df[fids]
                tmp_slice = cube_slice_df[cube_slice_df.columns[::-1]]
                tmp_slice.columns = self.years
                cube_dict[fids[0]] = tmp_slice
                pass
            the_cube = pd.Panel(cube_dict)
            if self.set_approp_cube(sheet, the_cube) == False:
                self.set_derived_cube(sheet, the_cube)
            pass
        pass

    def create_fid_collection(self, fids, years):
        num_years = len(years)
        fid_collection = [["" for x in range(num_years)] for y in range(len(fids))]
        for fid, row in zip(fids, range(len(fids))):
            tmp = fid[0]
            if tmp == None or tmp == u'' or tmp == u' ':
                continue
            for col in range(num_years):
                if col == 0:
                    fid_collection[row][col] = fid[0]
                else:
                    fid_collection[row][col] = fid[0] + "." + str(col)
        return fid_collection

    def get_years(self,the_df):
        tmp_years = []
        for i in the_df.iloc[1]:
            if pd.notnull(i):
                tmp_years.append(int(i))
        tmp_sorted_years = sorted(set(tmp_years))
        if len(self.years) == 0:
            self.years = tmp_sorted_years
        elif self.years != tmp_sorted_years:
            logging.fatal("Problem with Years: %s vs %s", self.years, tmp_sorted_years)
            exit()
        return

    def get_approp_df(self, sheet):
        the_df = None
        valid = True
        if sheet == "E07":
            the_df = self.E07_raw_df
        elif sheet == SI01_tag:
            the_df = self.SI01_raw_df
        elif sheet == "Assets":
            the_df = self.Assets_raw_df
        elif sheet == "CashFlow":
            the_df = self.CashFlow_raw_df
        else:
            valid = False
            logger.error("No sheet available in base class: %s", sheet)
        return (valid, the_df)

    def get_derived_df(self, sheet):
        the_df = None
        valid = False
        return (valid, the_df)

    def set_derived_cube(self, sheet, the_cube):
        return False

    def set_approp_cube(self, tag, the_cube):
        rv = True
        self.data_cubes[tag] = the_cube

        if tag == "E07":
            self.E07_cube = the_cube
        elif tag == SI01_tag:
            self.SI01_cube = the_cube
        elif tag == "Assets":
            self.Assets_cube = the_cube
        elif tag == "CashFlow":
            self.CashFlow_cube = the_cube
        else:
            logger.error("No cube for: %s", tag)
            self.set_derived_cube(tag, the_cube)
            rv = False
        return rv
    def get_df(self, co, tag):
        data_df = self.data_cubes[tag].major_xs(co).transpose()
        pct_df = data_df.pct_change(axis=1)
        pct_df = pct_df.drop(pct_df.columns[0],1)
        pct_labels = []
        data_labels = []
        for col in data_df.columns:
            pct_labels.append(str(col) + ".1")
            data_labels.append(str(col))
        pct_df.columns = pct_labels[0:len(pct_labels)-1]
        data_df.columns = data_labels
        the_df = pd.concat([data_df, pct_df], axis=1)
        the_df = the_df.reindex_axis(sorted(the_df.columns), axis=1)
        return the_df

    def get_template_column(self,sheet):
        pass

