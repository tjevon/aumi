from __future__ import print_function
import pandas as pd

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')
class BusinessType(object):
    """Base class for PC, Life, Health"""
    def __init__(self):
        self.BIG_raw_df = None

        self.cube_dict = {}

        self.years = []
        self.companies = set()
        pass

    def load_df(self, csv_filename):
        logger.info("Enter: %s", csv_filename)
        # read first 2 lines, determine what years are in play
        the_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode')
        self.get_years(the_df)

        # read df with fid for column label and group/company # for row label
        the_df = pd.read_csv(csv_filename, header=3, index_col=0, dtype='unicode')
        the_df = the_df.drop(['AMB#'])
        the_df = the_df.drop('Unnamed: 1', 1)
        the_df = the_df.drop('Unnamed: 2', 1)
        for i in the_df.index:
            if pd.notnull(i):
                self.companies.add(i)

        # doubtful that this exists any longer but just in case
        for i in the_df.columns:
            if i.find('Calc') != -1:
                the_df[i].replace(regex=True,inplace=True,to_replace=r'%',value=r'')

        the_df = the_df.astype(float)
        logger.info("Leave")
        return the_df

    def process_csvs( self, data_dir, file_names ):
        logger.info("Enter")

        return_file_names = []
        for file_name in file_names:
            if any(s in file_name for s in COMMON_TEMPLATE_TAGS):
                csv_filename = data_dir + "\\" + file_name
                the_df = self.load_df(csv_filename)
                # TODO: should we check to make sure fids don't already exist?
                try:
                    self.BIG_raw_df = pd.concat([self.BIG_raw_df, the_df], axis=1)
                except NameError:
                    self.BIG_raw_df = the_df
            else:
                return_file_names.append(file_name)
        logger.info("Leave")
        return return_file_names

    def construct_select_data_cubes(self, template_wb, sheet_list):
        for sheet in sheet_list:
            df_dict = {}
            fid_dict = template_wb.get_full_fid_list(sheet,self)[1]
            fid_collection_dict = self.create_fid_collection(fid_dict, self.years)
            df_return = self.get_approp_df(sheet)
            if df_return[0] == False:
                logger.error("No df available in derived class for sheet: %s", sheet)
                continue
            the_df = df_return[1]
            comp_dict = {}
            for key, fids in fid_collection_dict.iteritems():
                if fids[0] == '' or fids[0] == 'XXX':
                    continue
                if fids[0].find('AI') != -1:
                    comp_dict[key] = fids
                    continue
                cube_slice_df = the_df[fids]
                tmp_slice = cube_slice_df[cube_slice_df.columns[::-1]]
#                tmp_slice = cube_slice_df
                tmp_slice.columns = self.years
                df_dict[fids[0]] = tmp_slice
            self.do_calculations(comp_dict, template_wb, sheet, fid_collection_dict, df_dict)
            the_cube = pd.Panel(df_dict)
            self.cube_dict[sheet] = the_cube

    def do_calculations(self,comp_dict, template_wb, sheet, fid_collection_dict, cube_dict):
        for key, fids in comp_dict.iteritems():
            cell = key.replace('A',CALC_COL)
            my_formula = template_wb.get_formula(sheet,cell)
            func_idx = my_formula.find('(')
            args_idx = my_formula.find(')')
            func_key = my_formula[:func_idx]
            args_str = my_formula[func_idx+1:args_idx]
            args = args_str.split(",")
            slice = func_dict[func_key](args, fid_collection_dict, cube_dict)
            cube_dict[fids[0]] = slice
            pass

    def create_fid_collection(self, fid_dict, years):
        num_years = len(years)
        fid_collection_dict = {}
        for key, fid in fid_dict.iteritems():
            next_entry = ["" for x in range(num_years)]
            tmp = fid
            if tmp == None or tmp == u'' or tmp == u' ':
                fid_collection_dict[key] = next_entry
                continue
            for col in range(num_years):
                if col == 0:
                    next_entry[col] = fid
                else:
                    next_entry[col] = fid + "." + str(col)
                fid_collection_dict[key] = next_entry
        return fid_collection_dict

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
        df_return = (False, None)
        if any(s in sheet for s in COMMON_TEMPLATE_TAGS):
            the_df = self.BIG_raw_df
            df_return = (True, the_df)
        else:
            logger.error("No df available in base class for sheet: %s", sheet)
            df_return = self.get_derived_df(sheet)
        return df_return

    def get_derived_df(self, sheet):
        the_df = None
        valid = False
        return (valid, the_df)

    def set_derived_cube(self, sheet, the_cube):
        return False

    def get_df_including_pcts(self, co, tag):
        data_df = self.cube_dict[tag].major_xs(co).transpose()
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
        the_df = the_df.reindex_axis(sorted(the_df.columns,reverse=False), axis=1)
        return the_df

    def get_template_column(self,sheet):
        pass

    def get_bt_tag(self):
        return None
