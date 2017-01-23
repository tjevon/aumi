from __future__ import print_function
import pandas as pd

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')

class BusinessType(object):
    """Base class for PC, Life, Health"""
    def __init__(self):
        self.assets_raw_df = None
        self.cf_raw_df     = None
        self.E07_raw_df = None
        self.E10_raw_df = None
        self.SI01_raw_df    = None
        self.SI05_07_raw_df  = None

        self.assets_cube = None
        self.cf_cube = None
        self.E07_cube    = None
        self.E10_cube    = None
        self.SI01_cube    = None
        self.SI05_07_cube    = None

        self.years = []
        self.companies = set()
        pass

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

    def load_df(self, csv_filename):
        logger.info("Enter: %s", csv_filename)
        the_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode')
        self.get_years(the_df)

        # should read df with field # for column lable and group # for row label
        the_df = pd.read_csv(csv_filename, header=3, index_col=0, dtype='unicode')
        for i in the_df.index:
            if (i.find("Unnamed") != -1) or (i.find(".") != -1) or (i.find("AMB#") != -1):
                continue
            if pd.notnull(i):
                self.companies.add(i)
        for i in the_df.columns:
            if i.find('Calc') != -1:
                the_df[i].replace(regex=True,inplace=True,to_replace=r'%',value=r'')

        # try placing this above the for and eliminating if/continue
        the_df = the_df.drop(['AMB#'])
        the_df = the_df.drop('Unnamed: 1', 1)
        the_df = the_df.drop('Unnamed: 2', 1)
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
                self.assets_raw_df = the_df
            elif file_name.find(CashFlow_files) != -1:
                self.cf_raw_df = the_df
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

    def get_approp_df(self, sheet):
        if sheet == "E07":
            the_df = self.E07_raw_df
        elif sheet == "SI01":
            the_df = self.SI01_raw_df
        return the_df

    def set_approp_cube(self, sheet, the_cube):
        if sheet == "E07":
            self.E07_cube = the_cube
        elif sheet == "SI01":
            self.SI01_cube = the_cube
        return

    def construct_data_cubes(self, template_wb):
        for sheet in TEMPLATE_SHEETS:
            cube_dict = {}
            my_template_sheet = template_wb.xl_wb.sheets(sheet)
            my_fids = my_template_sheet.range(self.get_template_column(sheet)).options(ndim=2).value
            fid_collection = self.create_fid_collection(my_fids, self.years)
            the_df = self.get_approp_df(sheet)
            for fids in fid_collection:
                if fids[0] == '' or fids[0] == 'XXX':
                    continue
                cube_slice_df = the_df[fids]
                cube_dict[fids[0]] = cube_slice_df
                pass
            the_cube = pd.Panel(cube_dict)
            self.set_approp_cube(sheet, the_cube)
            pass
        pass

    def get_template_column(self,sheet):
        pass

    def create_fid_collection(self, fids, years):
        num_years = len(years)
        fid_collection = [["" for x in range(num_years)] for y in range(len(fids))]
        for fid, row in zip(fids, range(len(fids))):
            if fid[0] == None:
                continue
            for col in range(num_years):
                if col == 0:
                    fid_collection[row][col] = fid[0]
                else:
                    fid_collection[row][col] = fid[0] + "." + str(col)
        return fid_collection