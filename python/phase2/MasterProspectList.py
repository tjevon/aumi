from __future__ import print_function
import os

import PcBusinessType as pc
import LifeBusinessType as life
import HealthBusinessType as health
import TearSheetGenerator as tsg

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')


class MasterProspectList:
    """Class containing all PC, Life and Health classes,,, pretty much all the data"""
    def __init__(self, template_obj, target_obj, data_dir):
        self.file_dict = None
        self.template_obj = template_obj
        self.target_obj = target_obj
        self.business_types = {PC_tag:None, LIFE_tag:None, HEALTH_tag:None}
        self.build_business_types(data_dir)
        pass

    def get_filenames(self, data_dir):
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
        csv_data_dir = data_dir + "\\10_year_data"
        file_dict = self.get_file_dict(csv_data_dir)
        for entity_id in file_dict:
            if entity_id == PC_tag:
                self.business_types[entity_id] = pc.PcBusinessType()
            elif entity_id == LIFE_tag:
                self.business_types[entity_id] = life.LifeBusinessType()
            elif entity_id == HEALTH_tag:
                self.business_types[entity_id] = health.HealthBusinessType()
            else:
                continue
            self.business_types[entity_id].convert_csvs_to_raw_df(csv_data_dir, file_dict[entity_id])
            self.business_types[entity_id].construct_data_cube(self.template_obj)
        pass

    def get_companies(self, tag):
        if self.business_types[tag] is not None:
            return set(self.business_types[tag].companies)
        else:
            return None

    def build_specialty_cube(self):
        fids_for_cube = []
        for bus_type in BUSINESS_TYPES:
            bt = self.business_types[bus_type]
            if bt is None:
                continue
            section_tags = BUSINESS_TYPE_TAGS[bus_type]
            for tag in section_tags:
                template_fids_with_spaces = self.template_obj.get_display_fid_list(tag, bt)
                just_fids = filter(lambda a: a is not None, template_fids_with_spaces)
                just_fids = filter(lambda a: a != 'XXX', just_fids)
                fids_for_cube += just_fids
            cube = bt.data_cube[fids_for_cube]
            years = bt.years[::-1]
            cube = cube.loc[:,:,years]
            output_filename = "%s.csv" % bus_type
            for key in cube.minor_axis:
                df = cube.minor_xs(key)
                df.to_csv(output_filename, mode='a', encoding='utf-8')

                pass

            pass

    def build_trio_scorecard(self):
        logger.info("Enter")
        plot_sheets = [PC_tag, LIFE_tag, HEALTH_tag ]
        self.target_obj.add_xlsheets(plot_sheets)

        work_sheet = self.target_obj.get_target_worksheet(PC_tag)
        bus_type = self.business_types[PC_tag]
        cube = bus_type.data_cube

        years_of_interest = bus_type.years[::-1]

        inv_assets_fid = 'SF00025'  # dollar value of invested assets and cash

        # create df that contains invested assets for all companies for all years we care about
        inv_assets_df = cube[inv_assets_fid][years_of_interest]
        # create df that contains only companies with between 2 and 3 billion dollars
        filter_level = 20
        filter_num = filter_level * 1000000
        filter_df = inv_assets_df.loc[inv_assets_df[2015] > filter_num]
#        filter_df = filter_df.loc[inv_assets_df[2015] < 10000000]

        equity_fid = 'BI0001'  # percent of bonds / invested assets and cash
        equity_df = cube[equity_fid][years_of_interest]

        bonds_fid = 'AI0020'  # percent of bonds / invested assets and cash
        bonds_df = cube[bonds_fid][years_of_interest]

        # create df that contains percent bond/inv assets with only companies with 2-3 billion in invested assets

        equity_df = equity_df.loc[filter_df.index]
#        equity_df = equity_df[equity_df.columns[::-1]]
        equity_df = equity_df.transpose()

        bonds_df = bonds_df.loc[filter_df.index]
#        bonds_df = bonds_df[bonds_df.columns[::-1]]
        bonds_df = bonds_df.transpose()

        # calc median and mean for all companies
        median = bonds_df.median()
        mean = bonds_df.mean()

        title_str = "Bonds/Invested Assets (%d billion < Companies)" % filter_level
        fig = plt.figure()
        plt.plot(bonds_df)
        plt.title(title_str)
        plt.axis([years_of_interest[0], years_of_interest[-1], 0.0, 1.0])
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
        plt.axis([years_of_interest[0], years_of_interest[-1], 0.0, 1.0])
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
        # plt.axis([years_of_interest[0], years_of_interest[-1], 0.0, 1.0])
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
        # plt.axis([years_of_interest[0], years_of_interest[-1], 0.0, 1.0])
        # plt.ylabel('Percent')
        # plt.xlabel('Years')
        #
        # equity_plot = work_sheet.pictures.add(fig, name='Equities', update=True)
        # equity_plot.top = 0
        # equity_plot.left = 950
        # equity_plot.height = 200
        # equity_plot.width = 300

        interest_list = ['SF14450', 'AI04020', 'SF14453', 'AI04021', 'AI04022', 'AI04023', 'AI04024', 'SF14467', 'AI04025', 'AI04026', 'AI04027', 'AI04028', 'AI04029', 'AI04030' ]
        asset_interest_list = ['AI0020', 'BI0001', 'BI0002', 'BI0003', 'BI0004', 'BI0005',
                                'BI0006', 'BI0007', 'BI0008', 'BI0009', 'BI0010' ]

        sht_top = 0
        sht_left = 50
        for count, co in zip(range(0,len(bonds_df.columns)),bonds_df.columns):
            if (count % 5) == 0:
                sht_top += 200
                sht_left = 50
            company_id = co
            company_df = cube.major_xs(company_id)[asset_interest_list].transpose()
#            company_df = company_df[company_df.columns[::-1]]
            company_df = company_df[years_of_interest].transpose()

            title_str = "Asset Allocation (%s)" % co
            fig = plt.figure()
            plt.plot(company_df)
            plt.title(title_str)
            plt.axis([years_of_interest[0], years_of_interest[-1], 0.0, 1.0])
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

    def get_candidates(self, tag):
        companies = self.get_companies(tag)
        if companies is None:
            return None
        else:
            return self.get_arbitrary_subset(companies)
#            return None

    def get_arbitrary_subset(self, companies):
        """ Testing purposes only """
        tmp_set = set()
        for i in range(0, 1):
            tmp_set.add(companies.pop())
        return tmp_set

    def get_data_from_db(self):
        pass
