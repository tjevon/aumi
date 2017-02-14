from __future__ import print_function

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')


class Section(object):
    def __init__(self, tag, fid_collection_dict, comp_dict_ai, comp_dict_bi, comp_dict_ci):
        self.tag = tag
        self.fid_collection_dict = fid_collection_dict
        self.comp_dict_ai = comp_dict_ai
        self.comp_dict_bi = comp_dict_bi
        self.comp_dict_ci = comp_dict_ci
    pass


class BusinessType(object):
    """Base class for PC, Life, Health"""
    def __init__(self):
        self.section_map = {}
        self.years = []
        self.companies = set()
        self.raw_df = None
        self.data_cube = None
        pass

    def load_df(self, csv_filename):
        logger.info("Enter: %s", csv_filename)
        # read first 2 lines, determine what years are in play
        rv_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode')
        self.get_years(rv_df)

        # read df with fid for column label and group/company # for row label
        rv_df = pd.read_csv(csv_filename, header=3, index_col=0, dtype='unicode')
        rv_df = rv_df.drop(['AMB#'])
        rv_df = rv_df.drop('Unnamed: 1', 1)
        rv_df = rv_df.drop('Unnamed: 2', 1)
        for i in rv_df.index:
            if pd.notnull(i):
                self.companies.add(i)

        # doubtful that this exists any longer but just in case
        for i in rv_df.columns:
            if i.find('Calc') != -1:
                rv_df[i].replace(regex=True, inplace=True, to_replace=r'%', value=r'')

        rv_df = rv_df.astype(float)
        logger.info("Leave")
        return rv_df

    def convert_csvs_to_raw_df(self, data_dir, file_names):
        logger.info("Enter")
        return_file_names = []
        for name in file_names:
            if any(s in name for s in COMMON_TEMPLATE_TAGS):
                csv_filename = data_dir + "\\" + name
                df = self.load_df(csv_filename)
                # TODO: should we check to make sure fids don't already exist?
                try:
                    self.raw_df = pd.concat([self.raw_df, df], axis=1)
                except NameError:
                    self.raw_df = df
            else:
                return_file_names.append(name)
        logger.info("Leave")
        return return_file_names

    def construct_my_data_cube(self, template_wb, tag_list):
        logger.info("Enter")
        for tag in tag_list:
            comp_dict_ai = {}
            comp_dict_bi = {}
            comp_dict_ci = {}
            df_dict = {}
            fid_dict = template_wb.get_full_fid_list(tag, self)[1]
            fid_collection_dict = self.create_fid_collection(fid_dict, self.years)
            for key, fids in fid_collection_dict.iteritems():
                if fids[0] == '' or fids[0] == 'XXX':
                    continue
                if fids[0].find('AI') != -1:
                    comp_dict_ai[key] = fids
                    continue
                if fids[0].find('BI') != -1:
                    comp_dict_bi[key] = fids
                    continue
                if fids[0].find('CI') != -1:
                    comp_dict_ci[key] = fids
                    continue
                cube_slice_df = self.raw_df[fids]
                cube_slice_df.columns = self.years
                df_dict[fids[0]] = cube_slice_df
            cube = pd.Panel(df_dict)
            section = Section(tag, fid_collection_dict, comp_dict_ai, comp_dict_bi, comp_dict_ci)
            self.section_map[tag] = section
            if self.data_cube is None:
                self.data_cube = cube.copy()
            else:
                cube_list = [self.data_cube, cube]
                self.data_cube = pd.concat(cube_list, axis=0)
        self.data_cube = do_all_calculations(template_wb, self.section_map, self.data_cube)
        self.data_cube = self.calc_pct_change(self.data_cube)
        logger.info("Leave")
        return

    def create_fid_collection(self, fid_dict, years):
        num_years = len(years)
        fid_collection_dict = {}
        for key, fid in fid_dict.iteritems():
            next_entry = ["" for x in range(num_years)]
            tmp = fid
            if tmp is None or tmp == u'' or tmp == u' ':
                fid_collection_dict[key] = next_entry
                continue
            for col in range(num_years):
                if col == 0:
                    next_entry[col] = fid
                else:
                    next_entry[col] = fid + "." + str(col)
                fid_collection_dict[key] = next_entry
        return fid_collection_dict

    def calc_pct_change(self, cube):
        df_dict = {}
        for co in self.companies:
            company_df = cube.major_xs(co).transpose()
            company_df = company_df[company_df.columns[::-1]]  # reverse
            pct_df = company_df.pct_change(axis=1)
            pct_df = pct_df.drop(pct_df.columns[0], 1)  # drop nan's

            period = len(self.years) - 1
            pct_n_yr_df = company_df.pct_change(axis=1, periods=period)
            pct_n_yr_df = pct_n_yr_df.drop(pct_n_yr_df.columns[0:period], 1)  # drop nan's

            increasing_years = list(reversed(self.years))
            pct_labels = []
            for idx in range(0, len(self.years)-1):
                pct_labels.append(str(increasing_years[idx+1]) + "." + str(increasing_years[idx]))

            pct_labels.append(str(increasing_years[-1]) + ".1")
            n_yr_label = str(str(increasing_years[-1]) + "." +
                             str(increasing_years[0]))

            pct_df.columns = pct_labels[0:len(pct_labels)-1]
            pct_df = pct_df[pct_df.columns[::-1]]  # reverse

            pct_n_yr_df.columns = [n_yr_label]
            pct_df = pd.concat([pct_df, pct_n_yr_df], axis=1)

            company_df = company_df[company_df.columns[::-1]]  # reverse
            df = pd.concat([company_df, pct_df], axis=1)
            df_dict[co] = df
            pass
        rv_cube = pd.Panel(df_dict)
        rv_cube = rv_cube.swapaxes(0, 1)
        return rv_cube

    def get_df_including_pcts(self, co, fids):
        rv_df = self.data_cube.major_xs(co).transpose()
        rv_df = rv_df.loc[fids, :]
        rv_df = rv_df.reindex(fids)
        return rv_df

    def get_template_column(self, tag):
        pass

    def get_years(self, df):
        tmp_years = []
        for i in df.iloc[1]:
            if pd.notnull(i):
                tmp_years.append(int(i))
        tmp_sorted_years = sorted(set(tmp_years), reverse=True)
        if len(self.years) == 0:
            self.years = tmp_sorted_years
        return

    def get_bt_tag(self):
        return None
