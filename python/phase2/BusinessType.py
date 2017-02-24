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

    def __init__(self, bt_tag, period_idx, common_tag_list, bt_tag_list):
        self.bt_tag = bt_tag
        self.period_idx = period_idx
        self.common_tag_list = common_tag_list
        self.bt_tag_list = bt_tag_list
        self.complete_tag_list = self.common_tag_list + self.bt_tag_list

        self.periods = []
        self.companies = set()
        self.section_map = {}
        self.raw_df = None
        self.data_cube = None

    def construct_data_cube(self, data_dir, file_names, template_wb):
        self.convert_common_csvs_to_raw_df(data_dir, file_names, self.complete_tag_list)
        self.construct_my_data_cube(template_wb, self.complete_tag_list, self.period_idx)
        return

    def get_bt_tag(self):
        return self.bt_tag

    def load_df(self, csv_filename):
        logger.info("Enter: %s", csv_filename)
        # read first 2 lines, determine what periods are in play
        rv_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode')
        self.get_periods(rv_df)

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

    def convert_common_csvs_to_raw_df(self, data_dir, file_names, common_tag_list):
        logger.info("Enter")
        return_file_names = []
        for name in file_names:
            if any(s in name for s in common_tag_list):
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

    def construct_my_data_cube(self, template_wb, tag_list, y_or_q):
        logger.info("Enter")
        for tag in tag_list:
            comp_dict_ai = {}
            comp_dict_bi = {}
            comp_dict_ci = {}
            df_dict = {}
            fid_dict = template_wb.get_full_fid_list(tag, self, y_or_q)[1]
            fid_collection_dict = self.create_fid_collection(fid_dict, self.periods, y_or_q)
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
                cube_slice_df.columns = self.periods
                df_dict[fids[0]] = cube_slice_df
            cube = pd.Panel(df_dict)
            section = Section(tag, fid_collection_dict, comp_dict_ai, comp_dict_bi, comp_dict_ci)
            self.section_map[tag] = section
            if self.data_cube is None:
                self.data_cube = cube.copy()
            else:
                cube_list = [self.data_cube, cube]
                self.data_cube = pd.concat(cube_list, axis=0)
        if y_or_q == YEARLY_IDX:
            self.data_cube = do_all_calculations(template_wb, self.section_map, self.data_cube)
        self.data_cube = self.calc_pct_change(self.data_cube)
        logger.info("Leave")
        return

    def create_fid_collection(self, fid_dict, periods, y_or_q):
        num_periods = len(periods)
        fid_collection_dict = {}
        for key, fid in fid_dict.iteritems():
            next_entry = ["" for x in range(num_periods)]
            tmp = fid
            if tmp is None or tmp == u'' or tmp == u' ':
                fid_collection_dict[key] = next_entry
                continue
            for col in range(num_periods):
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

            period = len(self.periods) - 1
            pct_n_yr_df = company_df.pct_change(axis=1, periods=period)
            pct_n_yr_df = pct_n_yr_df.drop(pct_n_yr_df.columns[0:period], 1)  # drop nan's

            increasing_periods = list(reversed(self.periods))
            pct_labels = []
            for idx in range(0, len(self.periods)-1):
                pct_labels.append(str(increasing_periods[idx+1]) + "." + str(increasing_periods[idx]))

            pct_labels.append(str(increasing_periods[-1]) + ".1")
            n_yr_label = str(str(increasing_periods[-1]) + "." +
                             str(increasing_periods[0]))

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

    def get_specialty_cube(self, fids):
        cube = self.data_cube[fids]
        return cube

    def get_df_including_pcts(self, co, fids):
        rv_df = self.data_cube.major_xs(co).transpose()
        rv_df = rv_df.loc[fids, :]
        rv_df = rv_df.reindex(fids)
        return rv_df

    def get_periods(self, df):
        tmp_periods = []
        for i in df.iloc[1]:
            if pd.notnull(i):
                tmp_periods.append(i)
        #tmp_sorted_periods = sorted(set(tmp_periods), reverse=True)
        tmp_sorted_periods = tmp_periods
        if len(self.periods) == 0:
            self.periods = tmp_sorted_periods
        return

