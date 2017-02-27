from __future__ import print_function
import os

from BusinessType import *

import matplotlib.pyplot as plt

from AMB_defines import *

logger = logging.getLogger('twolane')


class MasterProspectList:
    """Class containing all PC, Life and Health classes,,, pretty much all the data"""
    def __init__(self, template_obj, target_obj, data_dir):
        self.file_dict = None
        self.template_obj = template_obj
        self.target_obj = target_obj
        self.business_types = {PC_tag: {}, LIFE_tag: {}, HEALTH_tag: {}}
        self.build_business_types(data_dir)
        pass

    @staticmethod
    def get_filenames(data_dir):
        logger.info("Enter")
        logger.info("DATAPATH: %s", data_dir)

        files = os.listdir(data_dir)
        file_names = []
        for name in files:
            logger.info("file: %s", name)
            file_names.append(name)
        logger.info("Leave")
        return file_names

    def get_file_dict(self, data_dir):
        logger.info("Enter")
        file_names = self.get_filenames(data_dir)
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

    def build_business_types(self, data_dir):
        yearly_csv_data_dir = data_dir + "\\10_year_data"
        file_dict = self.get_file_dict(yearly_csv_data_dir)
        if PC_tag in file_dict:
            self.business_types[PC_tag][YEARLY_IDX] = BusinessType(PC_tag, YEARLY_IDX,
                                                                   COMMON_TEMPLATE_TAGS,
                                                                   PC_TEMPLATE_TAGS)
            self.business_types[PC_tag][YEARLY_IDX].construct_data_cube(yearly_csv_data_dir,
                                                                        file_dict[PC_tag],
                                                                        self.template_obj)
        if LIFE_tag in file_dict:
            self.business_types[LIFE_tag][YEARLY_IDX] = BusinessType(LIFE_tag, YEARLY_IDX,
                                                                     COMMON_TEMPLATE_TAGS,
                                                                     LIFE_TEMPLATE_TAGS)
            self.business_types[LIFE_tag][YEARLY_IDX].construct_data_cube(yearly_csv_data_dir,
                                                                          file_dict[LIFE_tag],
                                                                          self.template_obj)
        if HEALTH_tag in file_dict:
            self.business_types[HEALTH_tag][YEARLY_IDX] = BusinessType(HEALTH_tag, YEARLY_IDX,
                                                                       COMMON_TEMPLATE_TAGS,
                                                                       HEALTH_TEMPLATE_TAGS)
            self.business_types[HEALTH_tag][YEARLY_IDX].construct_data_cube(yearly_csv_data_dir,
                                                                            file_dict[HEALTH_tag],
                                                                            self.template_obj)

        quarterly_csv_data_dir = data_dir + "\\2016_Q"
        file_dict = self.get_file_dict(quarterly_csv_data_dir)
        if PC_tag in file_dict:
            self.business_types[PC_tag][QUARTERLY_IDX] = BusinessType(PC_tag, QUARTERLY_IDX,
                                                                      COMMON_QUARTERLY_TAGS,
                                                                      PC_QUARTERLY_TAGS)
            self.business_types[PC_tag][QUARTERLY_IDX].construct_data_cube(quarterly_csv_data_dir,
                                                                           file_dict[PC_tag],
                                                                           self.template_obj)
        if LIFE_tag in file_dict:
            self.business_types[LIFE_tag][QUARTERLY_IDX] = BusinessType(LIFE_tag, QUARTERLY_IDX,
                                                                        COMMON_QUARTERLY_TAGS,
                                                                        LIFE_QUARTERLY_TAGS)
            self.business_types[LIFE_tag][QUARTERLY_IDX].construct_data_cube(quarterly_csv_data_dir,
                                                                             file_dict[LIFE_tag],
                                                                             self.template_obj)
        if HEALTH_tag in file_dict:
            self.business_types[HEALTH_tag][QUARTERLY_IDX] = BusinessType(HEALTH_tag, QUARTERLY_IDX,
                                                                          COMMON_QUARTERLY_TAGS,
                                                                          HEALTH_QUARTERLY_TAGS)
            self.business_types[HEALTH_tag][QUARTERLY_IDX].construct_data_cube(
                quarterly_csv_data_dir, file_dict[HEALTH_tag], self.template_obj)
        return

    def build_specialty_cube(self):
        fids_for_cube = []
        for bus_type in BUSINESS_TYPES:
            bt = self.business_types[bus_type]  # type: BusinessType
            if bt is None:
                continue
            section_tags = BUSINESS_TYPE_TAGS[bus_type]
            for tag in section_tags:
                template_fids_with_spaces = self.template_obj.get_display_fid_list(tag, bt)
                just_fids = filter(lambda a: a is not None, template_fids_with_spaces)
                just_fids = filter(lambda a: a != 'XXX', just_fids)
                fids_for_cube += just_fids
            cube = bt.data_cube[fids_for_cube]
            periods = bt.periods[::-1]
            cube = cube.loc[:, :, periods]
            output_filename = "%s.csv" % bus_type
            for key in cube.minor_axis:
                df = cube.minor_xs(key)
                df.to_csv(output_filename, mode='a', encoding='utf-8')
        return

    def build_trio_scorecard(self):
        logger.info("Enter")
        plot_sheets = [PC_tag, LIFE_tag, HEALTH_tag]
        self.target_obj.add_xlsheets(plot_sheets, YEARLY_IDX)

        work_sheet = self.target_obj.get_target_worksheet(PC_tag)
        bus_type = self.business_types[PC_tag][YEARLY_IDX]  # type: BusinessType
        cube = bus_type.data_cube

        tmp_list = bus_type.periods[::-1]
        periods_of_interest = [int(i) for i in tmp_list]

        inv_assets_fid = 'SF00025'  # dollar value of invested assets and cash

        # create df that contains invested assets for all companies for all periods we care about
        inv_assets_df = cube[inv_assets_fid][tmp_list]
        # create df that contains only companies with between 2 and 3 billion dollars
        filter_level = 20
        filter_num = filter_level * 1000000
        filter_df = inv_assets_df.loc[inv_assets_df['2015'] > filter_num]
#        filter_df = filter_df.loc[inv_assets_df[2015] < 10000000]

        equity_fid = 'BI0001'  # percent of bonds / invested assets and cash
        equity_df = cube[equity_fid][tmp_list]

        bonds_fid = 'AI0020'  # percent of bonds / invested assets and cash
        bonds_df = cube[bonds_fid][tmp_list]

        # create df that contains percent bond/inv assets with only companies with 2-3 billion in invested assets

        equity_df = equity_df.loc[filter_df.index]
#        equity_df = equity_df[equity_df.columns[::-1]]
        equity_df = equity_df.transpose()

        bonds_df = bonds_df.loc[filter_df.index]
#        bonds_df = bonds_df[bonds_df.columns[::-1]]
        bonds_df = bonds_df.transpose()

        # calc median and mean for all companies
#        median = bonds_df.median()
#        mean = bonds_df.mean()

        title_str = "Bonds/Invested Assets (%d billion < Companies)" % filter_level
        fig = plt.figure()
        plt.plot(bonds_df)
        plt.title(title_str)
        plt.axis([periods_of_interest[0], periods_of_interest[-1], 0.0, 1.0])
        plt.ylabel('Percent')
        plt.xlabel('Years')

        bonds_plot = work_sheet.pictures.add(fig, name='Bonds', update=True)
        bonds_plot.top = 0
        bonds_plot.left = 50
        bonds_plot.height = 200
        bonds_plot.width = 300

        title_str = "Equities/Invested Assets (%d billion < Companies)" % filter_level
        fig = plt.figure()
        plt.plot(equity_df)
        plt.title(title_str)
        plt.axis([periods_of_interest[0], periods_of_interest[-1], 0.0, 1.0])
        plt.ylabel('Percent')
        plt.xlabel('Years')

        equity_plot = work_sheet.pictures.add(fig, name='Equities', update=True)
        equity_plot.top = 0
        equity_plot.left = 350
        equity_plot.height = 200
        equity_plot.width = 300

        # df_median = bonds_df.loc[bonds_df[2015] < median[2015]]
        # df_median = df_median.transpose()
        #
        # df_mean = bonds_df.loc[bonds_df[2015] > mean[2015]]
        # df_mean = df_mean.transpose()

        # title_str = "Median - Bonds/Invested Assets (%d billion < Companies)" % filter_level
        # fig = plt.figure()
        # plt.plot(df_median)
        # plt.title(title_str)
        # plt.axis([periods_of_interest[0], periods_of_interest[-1], 0.0, 1.0])
        # plt.ylabel('Percent')
        # plt.xlabel('Years')
        #
        # equity_plot = work_sheet.pictures.add(fig, name='Equities', update=True)
        # equity_plot.top = 0
        # equity_plot.left = 650
        # equity_plot.height = 200
        # equity_plot.width = 300
        #
        # title_str = "Mean - Bonds/Invested Assets (%d billion < Companies)" % filter_level
        # fig = plt.figure()
        # plt.plot(df_mean)
        # plt.title(title_str)
        # plt.axis([periods_of_interest[0], periods_of_interest[-1], 0.0, 1.0])
        # plt.ylabel('Percent')
        # plt.xlabel('Years')
        #
        # equity_plot = work_sheet.pictures.add(fig, name='Equities', update=True)
        # equity_plot.top = 0
        # equity_plot.left = 950
        # equity_plot.height = 200
        # equity_plot.width = 300

        interest_list = ['SF14450', 'AI04020', 'SF14453', 'AI04021', 'AI04022', 'AI04023',
                         'AI04024', 'SF14467', 'AI04025', 'AI04026', 'AI04027', 'AI04028',
                         'AI04029', 'AI04030']
        asset_interest_list = ['AI0020', 'BI0001', 'BI0002', 'BI0003', 'BI0004', 'BI0005',
                               'BI0006', 'BI0007', 'BI0008', 'BI0009', 'BI0010']

        sht_top = 0
        sht_left = 50
        for count, co in zip(range(0, len(bonds_df.columns)), bonds_df.columns):
            if (count % 5) == 0:
                sht_top += 200
                sht_left = 50
            company_id = co
            company_df = cube.major_xs(company_id)[asset_interest_list].transpose()
#            company_df = company_df[company_df.columns[::-1]]
            company_df = company_df[tmp_list].transpose()

            title_str = "Asset Allocation (%s)" % co
            fig = plt.figure()
            plt.plot(company_df)
            plt.title(title_str)
            plt.axis([periods_of_interest[0], periods_of_interest[-1], 0.0, 1.0])
            plt.ylabel('Percent')
            plt.xlabel('Years')

            bonds_plot = work_sheet.pictures.add(fig, name=co, update=True)
            bonds_plot.top = sht_top
            bonds_plot.left = sht_left
            bonds_plot.height = 200
            bonds_plot.width = 300
            sht_left += 300

#        logger.info("%s in %s mean: %f median: %f", fid_of_interest, year_of_interest, mean, median)
        logger.info("Leave")
        return

    def get_companies(self, tag, y_or_q):
        if self.business_types[tag][y_or_q] is not None:
            return set(self.business_types[tag][y_or_q].companies)
        else:
            return None

    def get_candidates(self, tag, y_or_q):
        companies = self.get_companies(tag, y_or_q)
        if companies is None:
            return None
        else:
            return self.get_arbitrary_subset(companies)
#            return None

    @staticmethod
    def get_arbitrary_subset(companies):
        """ Testing purposes only """
        tmp_set = set()
        for i in range(0, 2):
            tmp_set.add(companies.pop())
        return tmp_set

    def get_data_from_db(self):
        pass
