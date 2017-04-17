from __future__ import print_function
import os

from PeriodType import *
from BusinessType import *

import matplotlib.pyplot as plt
import copy

from AMB_defines import *

logger = logging.getLogger('twolane')


class MasterProspectList:
    """Class containing all PC, Life and Health classes,,, pretty much all the data"""
    def __init__(self, template_obj, target_obj):
        self.file_dict = None
        self.template_obj = template_obj
        self.target_obj = target_obj
        self.business_types = {PC_tag: None, LIFE_tag: None, HEALTH_tag: None}
        self.build_business_types()
        self.yrly_proj_dict = {PC_tag: None, LIFE_tag: None, HEALTH_tag: None}
        self.qtrly_proj_dict = {PC_tag: None, LIFE_tag: None, HEALTH_tag: None}
        pass

    @staticmethod
    def get_arbitrary_subset(companies):
        """ Testing purposes only """
        tmp_dict = {}
        #        for i in range(0, len(companies)):
        count = 0
        for key, data in companies.iteritems():
            if count >= 1:
                break
            tmp_dict[key] = data
            count += 1
        return tmp_dict

    def build_business_types(self):
        self.business_types[PC_tag] = BusinessType(PC_tag, self.template_obj)
        #self.business_types[LIFE_tag] = BusinessType(LIFE_tag, self.template_obj)
        #self.business_types[HEALTH_tag] = BusinessType(HEALTH_tag, self.template_obj)
        return

    def build_projections_cube(self):
        logger.error("Enter")
        # first column, Liquid Assets
        for bt_tag, bt in self.business_types.items():
            if bt == None:
                continue
            projection_info = self.template_obj.get_projection_info(CashFlow_tag, bt_tag,
                                                                    YEARLY_IDX, CALCULATE_PROJECTION)
            cf_yrly_list = projection_info[0]

            proj_fid_list = self.template_obj.get_display_fid_list(Assets_tag, bt_tag, YEARLY_IDX, CALCULATE_PROJECTION)
            proj_fid_list += self.template_obj.get_display_fid_list(E10_tag, bt_tag, YEARLY_IDX, CALCULATE_PROJECTION)
            proj_fid_list += self.template_obj.get_display_fid_list(E07_tag, bt_tag, YEARLY_IDX, CALCULATE_PROJECTION)

            yrly_fid_list = copy.copy(proj_fid_list)
            yrly_fid_list += copy.copy(cf_yrly_list)

            yrly_fid_list = [x for x in yrly_fid_list if x != None]
            yrly_cube = bt.period_types[YEARLY_IDX].data_cube[yrly_fid_list]

            projection_info = self.template_obj.get_projection_info(CashFlow_tag, bt_tag,
                                                                    QUARTERLY_IDX, CALCULATE_PROJECTION)
            asset_to_cf_fid = defaultdict(list)
            for k, v in projection_info[1].iteritems():
                asset_to_cf_fid[v].append(k)

            cf_qtrly_list = projection_info[0]

            qtrly_cube = bt.period_types[QUARTERLY_IDX].data_cube[cf_qtrly_list]
            # get ride of pct columns
            q_size = qtrly_cube.shape[2]
            q_size /= 2
            q_size -= 1
            qtrly_cube = qtrly_cube.iloc[:,:,0:q_size]

            # TODO:  suspect in the future
            tmp_cube = yrly_cube[cf_yrly_list]
            fid_map = dict(zip(cf_yrly_list,cf_qtrly_list))
            # reset data fids to be the same as quarterly
            tmp_cube = tmp_cube.rename(fid_map)
            # grab the most recent year
            tmp_cube = tmp_cube[:,:,0:1]
            qtrly_cube = pd.concat([tmp_cube, qtrly_cube],axis=2)
            # grab most recent 4 quarters
            # qtrly_cube contains cumulative cash flow $ values at this point
            qtrly_cube = qtrly_cube[:,:,0:4]

            # TODO: highly suspect in future
            tmp_cube = yrly_cube[proj_fid_list]
            dollar_df = tmp_cube.minor_xs('2016')
            pct_df = tmp_cube.minor_xs('2016.2015')
            yrly_proj = dollar_df * pct_df
            self.yrly_proj_dict[bt_tag] = yrly_proj


            df_dict = {}
            for fid in cf_qtrly_list:
                slice = qtrly_cube[fid].astype(float)
                qtrly_matrix = np.diff(slice.values, axis=1)
                qtrly_matrix = qtrly_matrix * -1.0
                qtrly_df = pd.DataFrame(qtrly_matrix,index=slice.index)
                qtrly_df = pd.concat([qtrly_df, slice.iloc[:,-1]], axis=1)
                tsum = qtrly_df.sum(axis=1)
                qtrly_df = qtrly_df.div(tsum, axis=0)
                df_dict[fid] = qtrly_df
            pct_cube = pd.Panel(df_dict)

            pct_dict = defaultdict(list)
            projection_info = self.template_obj.get_projection_info(E10_tag, bt_tag,
                                                                    YEARLY_IDX, CALCULATE_PROJECTION)
            for k, v in projection_info[1].iteritems():
                pct_dict[v].append(k)

            projection_info = self.template_obj.get_projection_info(E07_tag, bt_tag,
                                                                    YEARLY_IDX, CALCULATE_PROJECTION)
            for k, v in projection_info[1].iteritems():
                pct_dict[v].append(k)

            try:
                proj_cube = None
                for k, v in asset_to_cf_fid.iteritems():
                    tmp_yrly_df = yrly_proj[pct_dict[k]]
                    tmp_pct_df = pct_cube[v[0]]
                    pro_dict = {}
                    for i in tmp_yrly_df.columns:
                        tmp_df = tmp_yrly_df[i]
                        pro_df = tmp_pct_df.mul(tmp_df,axis=0)
                        pro_dict[i] = pro_df
                    tmp_cube = pd.Panel(pro_dict)
                    proj_cube = pd.concat([proj_cube, tmp_cube],axis=0)
            except:
                pass
            self.qtrly_proj_dict[bt_tag] = proj_cube
            logger.error("Pct Cube created")

        logger.error("Leave")
        return

    def build_mpl_sheet(self):
        projection_sheets = ['MPL']
        self.target_obj.add_xlsheets(projection_sheets, '', MPL_IDX)
        sheet = self.target_obj.get_target_worksheet('MPL')
        line_no = 3
        for bt_tag, df in self.yrly_proj_dict.iteritems():
            try:
                bt = self.business_types[bt_tag]
                group_numbers = df.index.tolist()
                group_names = bt.get_group_names(group_numbers)
                group_names = [[i] for i in group_names]
                label_cell = 'A' + str(line_no)
                data_cell = 'B' + str(line_no)
                sheet.range(label_cell).value = group_names
                sheet.range(data_cell).options(index=False, header=False,).value = df
                line_no += df.shape[0]
            except:
                pass

    def build_specialty_cube(self):
        fids_for_cube = []
        for bus_type in BUSINESS_TYPES:
            bt = self.business_types[bus_type].period_types[YEARLY_IDX]  # type: PeriodType
            if bt is None:
                continue
            section_tags = BUSINESS_TYPE_TAGS[YEARLY_IDX][bus_type]
            for tag in section_tags:
                template_fids_with_spaces = self.template_obj.get_display_fid_list(tag,
                                                                                   bt.get_bt_tag(),
                                                                                   YEARLY_IDX,
                                                                                   DO_NOT_DISPLAY)
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
        self.target_obj.add_xlsheets(plot_sheets, '', YEARLY_IDX)

        work_sheet = self.target_obj.get_target_worksheet(PC_tag)
        bus_type = self.business_types[PC_tag].period_types[YEARLY_IDX]  # type: PeriodType
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
        if self.business_types[tag] is not None:
            if y_or_q in self.business_types[tag].period_types:
#                return set(self.business_types[tag].period_types[y_or_q].grp_unaf)
                company_info = self.business_types[tag].get_company_info(y_or_q)
                return company_info

        return None

    def get_candidates(self, tag, y_or_q):
        companies = self.get_companies(tag, y_or_q)
        if companies is None:
            return None
        else:
            return self.get_arbitrary_subset(companies)

    def get_data_from_db(self):
        pass
