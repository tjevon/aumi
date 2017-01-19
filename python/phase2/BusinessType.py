from __future__ import print_function
import pandas as pd

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')

class BusinessType(object):
    """Base class for PC, Life, Health"""
    def __init__(self):
        self.assets_df = None
        self.sch_d_df  = None
        self.cf_df     = None
        self.sch_ba_df = None
        self.sis_df    = None
        self.years = []
        self.companies = set()
        pass

    def load_df(self, csv_filename):
        logger.info("Enter: %s", csv_filename)
        tmp_years = []
        the_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode')
        for i in the_df.iloc[1]:
            if pd.notnull(i):
                tmp_years.append(int(i))
        tmp_sorted_years = sorted(set(tmp_years))
        if len(self.years) == 0:
            self.years = tmp_sorted_years
        elif self.years != tmp_sorted_years:
            logging.fatal("Problem with Years: %s vs %s", self.years, tmp_sorted_years)
            exit()

        # should read df with field # for column lable and group # for row label
        the_df = pd.read_csv(csv_filename, header=3, index_col=0, dtype='unicode')
        for i in the_df.index:
            if i.find("Unnamed") != -1:
                continue
            if i.find(".") != -1:
                continue
            if i.find("AMB#") != -1:
                continue
            if pd.notnull(i):
                self.companies.add(i)
        for i in the_df.columns:
            if i.find('Calc') != -1:
                the_df[i].replace(regex=True,inplace=True,to_replace=r'%',value=r'')

        logger.info("Leave")
        return the_df

    def process_csvs( self, data_dir, file_names ):
        logger.info("Enter")

        return_file_names = []
        for file_name in file_names:
            csv_filename = data_dir + "\\" + file_name
            the_df = self.load_df(csv_filename)

            if file_name.find(assets) != -1:
                self.assets_df = the_df
            elif file_name.find(cash_flow) != -1:
                self.cf_df = the_df
            elif file_name.find(sis) != -1:
                self.sis_df = the_df
            elif file_name.find(sch_d) != -1:
                self.sch_d_df = the_df
            elif file_name.find(sch_ba) != -1:
                self.sch_ba_df = the_df
            else:
                return_file_names.append(file_name)
        logger.info("Leave")
        return return_file_names
