from __future__ import print_function
from AMB_defines import *

logger = logging.getLogger('twolane')


class PeriodType(object):
    """Base class for PC, Life, Health"""

    def __init__(self, bt_obj, period_idx, template_obj):
        self.bt_tag = bt_obj.bt_tag
        self.period_idx = period_idx
        self.complete_tag_list = BUSINESS_TYPE_TAGS[self.period_idx][self.bt_tag]
        self.group_to_company = bt_obj.group_to_company

        if self.period_idx == YEARLY_IDX:
            # self.desired_periods = ['2016', '2015', '2014', '2013', '2012', '2011',
            #                         '2010', '2009', '2008', '2007', '2006']
            self.desired_periods = ['2016', '2015', '2014', '2013']
        self.avail_periods = []
        self.grp_unaf = set()
        self.section_map = {}
        self.raw_df = None
        self.data_cube = None
        self.position_cube = {}

        if self.period_idx == YEARLY_IDX:
            csv_data_dir = DATA_DIR + POSITIONS_DIR
            file_dict = PeriodType.get_file_dict(csv_data_dir)
            if self.bt_tag in file_dict:
                self.position_cube = self.build_position_cube(csv_data_dir, file_dict[self.bt_tag])

        if self.period_idx == YEARLY_IDX:
            csv_data_dir = DATA_DIR + YEARLY_DIR
        elif self.period_idx == QUARTERLY_IDX:
            csv_data_dir = DATA_DIR + QTRLY_DIR
        else:
            return
        file_dict = self.get_file_dict(csv_data_dir)
        if self.bt_tag in file_dict:
            self.construct_data_cube(csv_data_dir, file_dict[self.bt_tag], template_obj)
        return

    @staticmethod
    def get_filenames(csv_data_dir):
        logger.info("Enter")
        logger.info("DATAPATH: %s", csv_data_dir)
        files = os.listdir(csv_data_dir)
        logger.info("Leave")
        return files

    @staticmethod
    def create_fid_collection(fid_dict, avail_periods):
        num_periods = len(avail_periods)
        fid_collection_dict = {}
        for key, fid in fid_dict.items():
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

    @staticmethod
    def get_file_dict(csv_data_dir):
        logger.info("Enter")
        file_names = PeriodType.get_filenames(csv_data_dir)
        file_dict = {}

        for name in file_names:
            idx = name.find('_')
            if idx == -1:
                continue
            entity_id = name[:idx]
            if entity_id in file_dict:
                file_dict[entity_id].append(name)
            else:
                file_dict[entity_id] = [name]
        logger.info("Leave")
        return file_dict

    def construct_data_cube(self, csv_data_dir, file_names, template_wb):
        self.convert_common_csvs_to_raw_df(csv_data_dir, file_names, self.complete_tag_list)
        self.construct_my_data_cube(template_wb, self.complete_tag_list, self.period_idx)
        return

    def get_bt_tag(self):
        return self.bt_tag

    def load_df(self, csv_filename):
        logger.info("Enter: %s", csv_filename)
        # read first 2 lines, determine what periods are in play
        rv_df = pd.read_csv(csv_filename, header=0, index_col=0, nrows=2, dtype='unicode', thousands=",")
        self.get_periods(rv_df)

        # read df with fid for column label and group/company # for row label
        rv_df = pd.read_csv(csv_filename, header=3, index_col=0, dtype='unicode', thousands=",")
        rv_df = rv_df.drop(['AMB#'])
        rv_df = rv_df.drop('Unnamed: 1', 1)
        rv_df = rv_df.drop('Unnamed: 2', 1)
        for i in rv_df.index:
            if pd.notnull(i):
                self.grp_unaf.add(i)

        # doubtful that this exists any longer but just in case
        for i in rv_df.columns:
            if i.find('Calc') != -1:
                rv_df[i].replace(regex=True, inplace=True, to_replace=r'%', value=r'')

        rv_df = rv_df.astype(float)
        logger.info("Leave")
        return rv_df

    def convert_common_csvs_to_raw_df(self, csv_data_dir, file_names, complete_tag_list):
        logger.info("Enter")
        return_file_names = []
        for name in file_names:
            if any(s in name for s in complete_tag_list):
                csv_filename = csv_data_dir + "\\" + name
                df = self.load_df(csv_filename)
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
            comp_dict_vi = {}
            comp_dict_zi = {}
            df_dict = {}
            fid_dict = template_wb.get_full_fid_list(tag, self.get_bt_tag(), y_or_q)[1]
            fid_collection_dict = self.create_fid_collection(fid_dict, self.avail_periods)
            for key, fids in fid_collection_dict.items():
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
                if fids[0].find('ZI') != -1:
                    comp_dict_zi[key] = fids
                    continue
                if fids[0].find('VI') != -1:
                    comp_dict_vi[key] = fids
                    continue
                if fids[0] in df_dict:
                    logger.error("Second entry attempted %s", fids[0])
                    continue
                if self.data_cube is not None and fids[0] in self.data_cube.items:
                    logger.error("Second entry attempted %s", fids[0])
                    continue
                cube_slice_df = self.raw_df[fids]
                cube_slice_df.columns = self.avail_periods
                cube_slice_df = cube_slice_df[self.desired_periods]
                df_dict[fids[0]] = cube_slice_df

            cube = pd.Panel(df_dict)
            section = Section(tag, fid_collection_dict, comp_dict_ai, comp_dict_bi, comp_dict_ci,
                              comp_dict_vi, comp_dict_zi)
            self.section_map[tag] = section
            if self.data_cube is None:
                self.data_cube = cube.copy()
            else:
                cube_list = [self.data_cube, cube]
                self.data_cube = pd.concat(cube_list, axis=0)
        if y_or_q == YEARLY_IDX:
            self.data_cube = do_all_calculations(template_wb, self.section_map, self.data_cube,
                                                 self.position_cube, self.desired_periods)
        self.data_cube = self.calc_pct_change(self.data_cube)

        self.raw_df = None
        self.position_cube = None
        logger.info("Leave")

        return

    def reset_pct_labels(self, pct_df, pct_n_yr_df):
        increasing_periods = list(reversed(self.desired_periods))
        pct_labels = []
        for idx in range(0, len(self.desired_periods)-1):
            pct_labels.append(str(increasing_periods[idx+1]) + "." + str(increasing_periods[idx]))

        pct_labels.append(str(increasing_periods[-1]) + ".1")
        n_yr_label = str(str(increasing_periods[-1]) + "." +
                         str(increasing_periods[0]))
        pct_df.columns = pct_labels[0:len(pct_labels)-1]
        pct_n_yr_df.columns = [n_yr_label]
        return

    @staticmethod
    def get_pct_dfs(cube, co, period):
        company_df = cube.major_xs(co).transpose()
        company_df = company_df[company_df.columns[::-1]]  # reverse
        pct_df = company_df.pct_change(axis=1)
        pct_df = pct_df.drop(pct_df.columns[0], 1)  # drop nan's
        pct_n_yr_df = company_df.pct_change(axis=1, periods=period)
        pct_n_yr_df = pct_n_yr_df.drop(pct_n_yr_df.columns[0:period], 1)  # drop nan's
        return [company_df, pct_df, pct_n_yr_df]

    def calc_pct_change(self, cube):
        df_dict = {}
        period = len(self.desired_periods) - 1
        for co in self.grp_unaf:
            dfs = self.get_pct_dfs(cube, co, period)
            self.reset_pct_labels(dfs[1], dfs[2])

            dfs[1] = dfs[1][dfs[1].columns[::-1]]  # reverse
            dfs[0] = dfs[0][dfs[0].columns[::-1]]  # reverse

            dfs[1] = pd.concat([dfs[1], dfs[2]], axis=1)
            df = pd.concat([dfs[0], dfs[1]], axis=1)
            df_dict[co] = df
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
        tmp_sorted_periods = tmp_periods
        if len(self.avail_periods) == 0:
            self.avail_periods = tmp_sorted_periods
            if self.period_idx == QUARTERLY_IDX:
                self.desired_periods = self.avail_periods
        return

#    def set_group_to_company(self, grp_to_co):
#        self.group_to_company = { key: grp_to_co[key] for key in grp_to_co if key in self.grp_unaf}

#        return

    def build_df_dict(self, csv_filename, start, year, issuer_code, issuer_type, lineno):
        products_owned = {}
        df = pd.read_csv(csv_filename, header=start, dtype='unicode', thousands=",")
        df = df.ix[1:]
        df = df.rename(columns={'Unnamed: 0': 'AMB#'})
        df = df.drop('Unnamed: 1', 1)
        df = df.drop('Unnamed: 2', 1)
        df['AMB#'] = df['AMB#'].apply(lambda x: x.zfill(6))
        df['cusip'] = df[issuer_code].map(str) + df[issuer_type].map(str)
        df['cusip'] = df['cusip'].apply(lambda x: x.zfill(8))
        df['year'] = year
        df = df.loc[df[lineno] != '99999']

        empty_groups = 0
        for grp, companies in self.group_to_company.items():
            df_grp = df[df['AMB#'].isin(companies)]
            if df_grp.size == 0:
                empty_groups += 1
                logger.error("Empty group %s: ", grp)
                continue
            products_owned[grp] = df_grp
            pass
        logger.error("Empty group count %d: ", empty_groups)
        return products_owned

    def order_file_list(self, file_list):
        ba_files = [x for x in file_list if x.find('_BA_') != -1]
        file_list = [x for x in file_list if x not in ba_files]
        bonds_owned_list = [x for x in file_list if x.find('Bonds_Owned') != -1]
        bonds_owned_list = bonds_owned_list[::-1]
        stocks_owned_list = [x for x in file_list if x.find('Stocks_Owned') != -1]
        stocks_owned_list = stocks_owned_list[::-1]
        acq_list = [x for x in file_list if x.find('Acquired') != -1]
        disp_list = [x for x in file_list if x.find('Disposed') != -1]
        if self.file_checks(bonds_owned_list, stocks_owned_list, acq_list, disp_list) is False:
            logger.fatal("Please check to make sure all needed files are available")
            exit()
        return [bonds_owned_list, stocks_owned_list, acq_list, disp_list]

    @staticmethod
    def file_checks(bo_list, so_list, a_list, d_list):
        bo_list = sorted(bo_list)
        so_list = sorted(so_list)
        a_list = sorted(a_list)
        d_list = sorted(d_list)
        if len(bo_list) == 0:
            logger.error("Number of Bonds Owned files is zero")
            return False
        if len(bo_list) != len(so_list):
            logger.error("Number of Bonds Owned files %d differs from Stocks Owned files %d",
                         len(bo_list), len(so_list))
            return False
        if len(a_list) != len(d_list):
            logger.error("Number of Acquired files %d differs from Disposed files %d",
                         len(a_list), len(d_list))
            return False
        for i in range(0, len(bo_list)):
            idx = bo_list[i].find('.csv')
            if idx == -1:
                logger.error("Bonds Owned not a .csv file")
                return False
            bo_year = bo_list[i][idx-4:idx]

            idx = so_list[i].find('.csv')
            if idx == -1:
                logger.error("Stocks Owned not a .csv file")
                return False
            so_year = so_list[i][idx-4:idx]
            if bo_year != so_year:
                logger.error("Years are incorrect for Bonds and Stocks Owned")
                return False
        return True

    def build_position_cube(self, csv_data_dir, file_list):
        segment_list = self.order_file_list(file_list)

        year_dict = {}
        for b, s in zip(segment_list[0], segment_list[1]):
            if self.period_idx == YEARLY_IDX:
                if b.find("EOY") == -1:
                    continue
            idx = b.find('.csv')
            year = b[idx-4:idx]

            csv_filename = csv_data_dir + "\\" + b
            bonds_owned_dict = self.build_df_dict(csv_filename, 5, year, 'IB00019', 'IB00020', 'IB00050')
            csv_filename = csv_data_dir + "\\" + s
            stocks_owned_dict = self.build_df_dict(csv_filename, 6, year, 'IS00019', 'IS00020', 'IS00050')
            year_dict[year] = (bonds_owned_dict, stocks_owned_dict)
            pass
        return year_dict
