from __future__ import print_function

from BusinessType import *

import copy
import math
from itertools import chain

from AMB_defines import *

logger = logging.getLogger('twolane')


class MasterProspectList:
    """Class containing all PC, Life and Health classes,,, pretty much all the data"""
    def __init__(self, template_obj, target_obj, quarterly):
        self.file_dict = None
        self.quarterly_arg = quarterly
        self.template_obj = template_obj
        self.target_obj = target_obj
        self.business_types = {PC_tag: None, LIFE_tag: None, HEALTH_tag: None}
        self.build_business_types(self.quarterly_arg)
        self.yrly_proj_dict = {PC_tag: None, LIFE_tag: None, HEALTH_tag: None}
        self.qtrly_proj_dict = {PC_tag: None, LIFE_tag: None, HEALTH_tag: None}
        pass

    @staticmethod
    def get_arbitrary_subset(companies):
        """ Testing purposes only """
        tmp_dict = {}
        MAX_ITEMS = 1000
        count = 0
        for key, data in companies.items():
            if count > MAX_ITEMS-1:
                break
            if type(data) is float:
                continue
            tmp_dict[key] = data
            count += 1

        return tmp_dict

    def build_business_types(self, quarterly_arg):
        x = 1
        if x > 0:
            self.business_types[PC_tag] = BusinessType(PC_tag, self.template_obj, quarterly_arg)
        if x > 1:
            self.business_types[LIFE_tag] = BusinessType(LIFE_tag, self.template_obj, quarterly_arg)
        if x > 2:
            self.business_types[HEALTH_tag] = BusinessType(HEALTH_tag, self.template_obj, quarterly_arg)
        return

    @staticmethod
    def single_exp_smooth(tmp_cube, a, period_dim):
        proj_df = None
        for fid, df in tmp_cube.items():
            df2 = df.iloc[:, 0: period_dim]
            df3 = df2[df2.columns[:: -1]]
            proj = df3.ewm(alpha=a, axis=1)
            df4 = proj.mean()
            df5 = df4[df4.columns[::-1]]
            df2_p = df2.iloc[:, 0:1] * a
            df5_p = df5.iloc[:, 0:1] * (1.0 - a)
            df6 = df2_p + df5_p
            df7 = df6 - df2.iloc[:, 0:1]
            df7.columns = [fid]
            proj_df = pd.concat([proj_df, df7], axis=1)
        return proj_df

    @staticmethod
    def double_exp_smooth(tmp_cube, period_dim):
        a = 0.95
        g = 0.9
        proj_df = None
        # for each fid
        for fid, df in tmp_cube.iteritems():
            df2 = df.iloc[:, 0:period_dim]
            df3 = df2[df2.columns[::-1]]

            df_d = df3.copy(deep=True)
            df_s = df3.copy(deep=True)
            df_b = df3.copy(deep=True)
            df_b.iloc[:, 0] = df_b.iloc[:, 1] - df_b.iloc[:, 0]

            for i in range(df2.shape[1]):
                if i > 0:
                    s_prev = df_s.iloc[:, i-1]
                    b_prev = df_b.iloc[:, i-1]

                    df_s.iloc[:, i] = a * df3.iloc[:, i] + (1.0 - a) * (s_prev + b_prev)
                    df_b.iloc[:, i] = g * (df_s.iloc[:, i] - s_prev) + (1.0 - g) * b_prev

                    if i > 1:
                        df_d.iloc[:, i] = df_s.iloc[:, i] + df_b.iloc[:, i]
            df_d = df_d[df_d.columns[::-1]]
            df_st = df_d.iloc[:, 0:1] - df2.iloc[:, 0:1]
            df_st.columns = [fid]
            proj_df = pd.concat([proj_df, df_st], axis=1)
        return proj_df

    @staticmethod
    def previous_yr_pct(tmp_cube, most_recent_year):
        prev_year = int(most_recent_year) - 1
        most_recent_pct = most_recent_year + '.' + str(prev_year)

        dollar_df = tmp_cube.minor_xs(most_recent_year)
        pct_df = tmp_cube.minor_xs(most_recent_pct)
        yrly_proj = dollar_df * pct_df
        return yrly_proj

    def build_projections_cube(self):
        logger.error("Enter")
        # first column, Liquid Assets
        for bt_tag, bt in self.business_types.items():
            if bt is None:
                continue
            projection_info = self.template_obj.get_projection_info(CashFlow_tag, bt_tag,
                                                                    YEARLY_IDX, DISPLAY_MPL)
            cf_yrly_list = projection_info[0]

            proj_fid_list = self.template_obj.get_projection_info(Assets_tag, bt_tag, YEARLY_IDX, DISPLAY_MPL)[0]
            proj_fid_list += self.template_obj.get_projection_info(Liquid_Assets_tag, bt_tag, YEARLY_IDX, DISPLAY_MPL)[0]
            proj_fid_list += self.template_obj.get_projection_info(E07_tag, bt_tag, YEARLY_IDX, DISPLAY_MPL)[0]
            proj_fid_list += self.template_obj.get_projection_info(Real_Estate_tag, bt_tag, YEARLY_IDX, DISPLAY_MPL)[0]

            yrly_fid_list = copy.copy(proj_fid_list)
            yrly_fid_list += copy.copy(cf_yrly_list)

            yrly_fid_list = [x for x in yrly_fid_list if x is not None]
            yrly_cube = bt.period_types[YEARLY_IDX].data_cube[yrly_fid_list]

            projection_info = self.template_obj.get_projection_info(CashFlow_tag, bt_tag,
                                                                    QUARTERLY_IDX, DISPLAY_MPL)
            asset_to_cf_fid = defaultdict(list)
            for k, v in projection_info[1].items():
                asset_to_cf_fid[v].append(k)

            cf_qtrly_list = projection_info[0]

            qtrly_cube = bt.period_types[QUARTERLY_IDX].data_cube[cf_qtrly_list]
            # get ride of pct columns
            q_size = qtrly_cube.shape[2]
            # ignore deltas
            q_size /= 2
            tmp_cube = qtrly_cube.iloc[:, :, 0:q_size]
            q_slice = qtrly_cube.iloc[0,:,:]
            q_idx = q_slice.columns[0].find('20')
            q_year = q_slice.columns[0][q_idx:q_idx+4]

            y_slice = yrly_cube.iloc[0,:,:]
            y_year = str(y_slice.columns[0])

            if y_year == q_year:
                # special case for EOY because we don't have 4th quarter

                q_size -= 1
                qtrly_cube = qtrly_cube.iloc[:, :, 0:q_size]

                # TODO:  suspect in the future
                tmp_cube = yrly_cube[cf_yrly_list]
                fid_map = dict(zip(cf_yrly_list, cf_qtrly_list))
                # reset data fids to be the same as quarterly
                tmp_cube = tmp_cube.rename(fid_map)
                # grab the most recent year
                tmp_cube = tmp_cube[:, :, 0:1]
                qtrly_cube = pd.concat([tmp_cube, qtrly_cube], axis=2)
                # grab most recent 4 quarters
                # qtrly_cube contains cumulative cash flow $ values at this point
                qtrly_cube = qtrly_cube[:, :, 0:4]
            else:
                # quarterlies are newer than most recent yearly
                # following is garbage
                q_size -= 1
                qtrly_cube = qtrly_cube.iloc[:, :, 0:q_size]

                # TODO:  suspect in the future
                tmp_cube = yrly_cube[cf_yrly_list]
                fid_map = dict(zip(cf_yrly_list, cf_qtrly_list))
                # reset data fids to be the same as quarterly
                tmp_cube = tmp_cube.rename(fid_map)
                # grab the most recent year
                tmp_cube = tmp_cube[:, :, 0:1]
                qtrly_cube = pd.concat([tmp_cube, qtrly_cube], axis=2)
                # grab most recent 4 quarters
                # qtrly_cube contains cumulative cash flow $ values at this point
                qtrly_cube = qtrly_cube[:, :, 0:4]

            # TODO: highly suspect in future
            most_recent_year = y_year
            tmp_cube = yrly_cube[proj_fid_list]
            a = 0.9
            period_dim = tmp_cube.shape[2]/2
            yrly_proj = None
            which = MODEL_DBL_EXP
            if which == MODEL_PREV_YEAR:
                yrly_proj = self.previous_yr_pct(tmp_cube, most_recent_year)
            elif which == MODEL_SGL_EXP:
                yrly_proj = self.single_exp_smooth(tmp_cube, a, period_dim)
            elif which == MODEL_DBL_EXP:
                yrly_proj = self.double_exp_smooth(tmp_cube, period_dim)

            self.yrly_proj_dict[bt_tag] = yrly_proj

            # following takes the quarterly cumulative values, determines incremental 1/4ly
            # increases, then calculates the percent change for previous 4 1/4s
            # Assumptions are EOY is most recent, values decrease relative to same year
            df_dict = {}
            for fid in cf_qtrly_list:
                q_slice = qtrly_cube[fid].astype(float)
                qtrly_matrix = np.diff(q_slice.values, axis=1)
                qtrly_matrix = qtrly_matrix * -1.0
                qtrly_df = pd.DataFrame(qtrly_matrix, index=q_slice.index)
                qtrly_df = pd.concat([qtrly_df, q_slice.iloc[:, -1]], axis=1)
                tsum = qtrly_df.sum(axis=1)
                qtrly_df = qtrly_df.div(tsum, axis=0)
                df_dict[fid] = qtrly_df
            pct_cube = pd.Panel(df_dict)

            pct_dict = defaultdict(list)
            projection_info = self.template_obj.get_projection_info(Liquid_Assets_tag, bt_tag,
                                                                    YEARLY_IDX,
                                                                    DISPLAY_MPL)
            for k, v in projection_info[1].items():
                pct_dict[v].append(k)

            projection_info = self.template_obj.get_projection_info(E07_tag, bt_tag,
                                                                    YEARLY_IDX,
                                                                    DISPLAY_MPL)
            for k, v in projection_info[1].items():
                pct_dict[v].append(k)

            try:
                proj_cube = None
                for k, v in asset_to_cf_fid.items():
                    tmp_yrly_df = yrly_proj[pct_dict[k]]
                    tmp_pct_df = pct_cube[v[0]]
                    pro_dict = {}
                    for i in tmp_yrly_df.columns:
                        tmp_df = tmp_yrly_df[i]
                        pro_df = tmp_pct_df.mul(tmp_df, axis=0)
                        pro_dict[i] = pro_df
                    tmp_cube = pd.Panel(pro_dict)
                    proj_cube = pd.concat([proj_cube, tmp_cube], axis=0)
            except:
                continue
            self.qtrly_proj_dict[bt_tag] = proj_cube
            logger.error("Pct Cube created")

        logger.error("Leave")
        return

    def build_mpl_sheet(self, sheet_name, section_tag):
        projection_sheets = [sheet_name]
        self.target_obj.add_xlsheets(projection_sheets, '', MPL_IDX)
        sheet = self.target_obj.get_target_worksheet(projection_sheets[0])
        line_no = 4
        for bt_tag, proj_df in self.yrly_proj_dict.items():
            try:
                bt = self.business_types[bt_tag]
                proj_fid_list = self.template_obj.get_projection_info(Assets_tag, bt_tag,
                                                                       YEARLY_IDX,
                                                                       DISPLAY_MPL)[0]
                asset_col_size = len(proj_fid_list)
                proj_fid_list += self.template_obj.get_projection_info(section_tag, bt_tag,
                                                                       YEARLY_IDX,
                                                                       DISPLAY_MPL)[0]
                if section_tag == E07_tag:
                    proj_fid_list += self.template_obj.get_projection_info(Real_Estate_tag,
                                                                           bt_tag, YEARLY_IDX,
                                                                           DISPLAY_MPL)[0]

                proj_df = proj_df[proj_fid_list]
                yrly_data_cube = bt.period_types[YEARLY_IDX].data_cube[proj_fid_list]
                cur_yr_df = yrly_data_cube.minor_xs('2016')

                group_numbers = proj_df.index.tolist()
                cur_yr_df = cur_yr_df.reindex(group_numbers)

                group_info_df = bt.get_group_info_df(group_numbers)

                cur_yr_df.columns = [str(c) + '_y' for c in cur_yr_df.columns]
                tmp_columns = zip(cur_yr_df.columns[asset_col_size:], proj_df.columns[asset_col_size:])
                tmp_columns = list(chain.from_iterable(tmp_columns))
                tmp_df = pd.concat([cur_yr_df.iloc[:, asset_col_size:], proj_df.iloc[:, asset_col_size:]],axis=1)
                tmp_df = tmp_df[tmp_columns]

                group_info_df["Industry Sector"] = bt_tag
                label_cell = 'A' + str(line_no)
                asset_cell = 'E' + str(line_no)
                data_cell = 'R' + str(line_no)
                sheet.range(label_cell).options(index=False, header=False).value = group_info_df
                sheet.range(asset_cell).options(index=False, header=False).value = cur_yr_df.iloc[:, 0:asset_col_size]
                sheet.range(data_cell).options(index=False, header=False).value = tmp_df
                line_no += proj_df.shape[0]
            except:
                logger.error("Exception caught:")
                pass
        return

    def build_industry_info_sheet(self, sheet_name, section_tag, bt_tag):
        projection_sheets = [sheet_name]
        self.target_obj.add_xlsheets(projection_sheets, '', IF_IDX)
        sheet = self.target_obj.get_target_worksheet(projection_sheets[0])
        line_no = 2
        if bt_tag in self.business_types:
            try:
                bt = self.business_types[bt_tag]
                proj_fid_list = self.template_obj.get_projection_info(Assets_tag, bt_tag,
                                                                       YEARLY_IDX,
                                                                       DISPLAY_MPL)[0]
                asset_size = len(proj_fid_list)
                proj_fid_list += self.template_obj.get_projection_info(Liquid_Assets_tag, bt_tag,
                                                                       YEARLY_IDX,
                                                                       DISPLAY_MPL)[0]
                liq_size = len(proj_fid_list) - asset_size
                proj_fid_list += self.template_obj.get_projection_info(E07_tag, bt_tag,
                                                                       YEARLY_IDX,
                                                                       DISPLAY_MPL)[0]
                e07_size = len(proj_fid_list) - liq_size - asset_size

                proj_fid_list += self.template_obj.get_projection_info(Real_Estate_tag, bt_tag,
                                                                       YEARLY_IDX,
                                                                       DISPLAY_MPL)[0]
                re_size = len(proj_fid_list) - e07_size - liq_size - asset_size
                cum_size = len(proj_fid_list)

                yrly_data_cube = bt.period_types[YEARLY_IDX].data_cube[proj_fid_list]
                prev_yr_df = yrly_data_cube.minor_xs('2015')
                cur_yr_df = yrly_data_cube.minor_xs('2016')

                cur_sum_df = cur_yr_df.sum(axis=0)
                prev_sum_df = prev_yr_df.sum(axis=0)
                delta_sum_df = cur_sum_df - prev_sum_df

                asset_mean_df = cur_yr_df.mean(axis=0)
                asset_median_df = cur_yr_df.median(axis=0)
                asset_first_q = cur_yr_df.quantile(q=0.25,axis=0)
                asset_third_q = cur_yr_df.quantile(q=0.75,axis=0)
                asset_sdev = cur_yr_df.std(axis=0)

                proj_fid_list = self.template_obj.get_projection_info(section_tag, bt_tag,
                                                                      YEARLY_IDX,
                                                                      DISPLAY_MPL)[0]
                asset_alloc_size = len(proj_fid_list)
                yrly_data_cube = bt.period_types[YEARLY_IDX].data_cube[proj_fid_list]

                # TODO: beware, quarterly and next year
                asset_alloc_df = yrly_data_cube.minor_xs('2016')

                mean_df = asset_alloc_df.mean(axis=0)
                median_df = asset_alloc_df.median(axis=0)
                first_q = asset_alloc_df.quantile(q=0.25,axis=0)
                third_q = asset_alloc_df.quantile(q=0.75,axis=0)
                sdev = asset_alloc_df.std(axis=0)

                line_no = 2
                DOLLAR_IDX = 7

                START_IDX = 0
                FINISH_IDX = DOLLAR_IDX
                data_cell = 'B' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = cur_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'C' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = prev_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'D' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = delta_sum_df[START_IDX:FINISH_IDX]

                data_cell = 'E' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = mean_df[START_IDX:FINISH_IDX]
                data_cell = 'F' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = median_df[START_IDX:FINISH_IDX]
                data_cell = 'G' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = first_q[START_IDX:FINISH_IDX]
                data_cell = 'H' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = third_q[START_IDX:FINISH_IDX]
                data_cell = 'I' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = sdev[START_IDX:FINISH_IDX]
                line_no = 9
                START_IDX = DOLLAR_IDX
                FINISH_IDX = asset_size
                data_cell = 'E' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = asset_mean_df[START_IDX:FINISH_IDX]
                data_cell = 'F' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = asset_median_df[START_IDX:FINISH_IDX]
                data_cell = 'G' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = asset_first_q[START_IDX:FINISH_IDX]
                data_cell = 'H' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = asset_third_q[START_IDX:FINISH_IDX]
                data_cell = 'I' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = asset_sdev[START_IDX:FINISH_IDX]

                line_no = 17
                START_IDX = asset_size
                FINISH_IDX = asset_size + liq_size
                data_cell = 'B' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = cur_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'C' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = prev_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'D' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = delta_sum_df[START_IDX:FINISH_IDX]

                data_cell = 'E' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = mean_df[START_IDX:FINISH_IDX]
                data_cell = 'F' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = median_df[START_IDX:FINISH_IDX]
                data_cell = 'G' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = first_q[START_IDX:FINISH_IDX]
                data_cell = 'H' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = third_q[START_IDX:FINISH_IDX]
                data_cell = 'I' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = sdev[START_IDX:FINISH_IDX]

                line_no += liq_size
                line_no += 2

                START_IDX = FINISH_IDX
                FINISH_IDX = asset_size + liq_size + e07_size
                data_cell = 'B' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = cur_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'C' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = prev_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'D' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = delta_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'E' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = mean_df[START_IDX:FINISH_IDX]
                data_cell = 'F' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = median_df[START_IDX:FINISH_IDX]
                data_cell = 'G' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = first_q[START_IDX:FINISH_IDX]
                data_cell = 'H' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = third_q[START_IDX:FINISH_IDX]
                data_cell = 'I' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = sdev[START_IDX:FINISH_IDX]

                line_no += e07_size
                line_no += 2

                START_IDX = FINISH_IDX
                FINISH_IDX = asset_size + liq_size + e07_size + re_size
                data_cell = 'B' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = cur_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'C' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = prev_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'D' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = delta_sum_df[START_IDX:FINISH_IDX]
                data_cell = 'E' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = mean_df[START_IDX:FINISH_IDX]
                data_cell = 'F' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = median_df[START_IDX:FINISH_IDX]
                data_cell = 'G' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = first_q[START_IDX:FINISH_IDX]
                data_cell = 'H' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = third_q[START_IDX:FINISH_IDX]
                data_cell = 'I' + str(line_no)
                sheet.range(data_cell).options(index=False, header=False,).value = sdev[START_IDX:FINISH_IDX]


                line_no = 2
                DOLLAR_IDX = 8
                cum_size = liq_size + asset_size

                # START_IDX = 0
                # FINISH_IDX = DOLLAR_IDX
                # data_cell = 'E' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = mean_df[START_IDX:FINISH_IDX]
                # data_cell = 'F' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = median_df[START_IDX:FINISH_IDX]
                # data_cell = 'G' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = first_q[START_IDX:FINISH_IDX]
                # data_cell = 'H' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = third_q[START_IDX:FINISH_IDX]
                # data_cell = 'I' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = sdev[START_IDX:FINISH_IDX]
                #
                # START_IDX = DOLLAR_IDX-1
                # FINISH_IDX = liq_size
                # data_cell = 'E' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = mean_df[START_IDX:FINISH_IDX]
                # data_cell = 'F' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = median_df[START_IDX:FINISH_IDX]
                # data_cell = 'G' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = first_q[START_IDX:FINISH_IDX]
                # data_cell = 'H' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = third_q[START_IDX:FINISH_IDX]
                # data_cell = 'I' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = sdev[START_IDX:FINISH_IDX]
                #
                # ####
                # data_cell = 'E' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = mean_df[liq_size:asset_alloc_size]
                # data_cell = 'F' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = median_df[liq_size:asset_alloc_size]
                # data_cell = 'G' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = first_q[liq_size:asset_alloc_size]
                # data_cell = 'H' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = third_q[liq_size:asset_alloc_size]
                # data_cell = 'I' + str(line_no)
                # sheet.range(data_cell).options(index=False, header=False,).value = sdev[liq_size:asset_alloc_size]

#                asset_cell = 'B' + str(line_no)
#                sheet.range(asset_cell).options(index=False, header=False,).value = cur_yr_df.iloc[:, 0:5]
#                line_no += 5
            except:
                pass
        return

    def build_specialty_cube(self):
        fids_for_cube = []
        for bus_type in BUSINESS_TYPES:
            pt = self.business_types[bus_type].period_types[YEARLY_IDX]  # type: PeriodType
            if pt is None:
                continue
            section_tags = BUSINESS_TYPE_TAGS[YEARLY_IDX][bus_type]
            for tag in section_tags:
                template_fids_with_spaces = self.template_obj.get_display_fid_list(tag,
                                                                                   pt.get_bt_tag(),
                                                                                   YEARLY_IDX,
                                                                                   DO_NOT_DISPLAY)
                just_fids = filter(lambda a: a is not None, template_fids_with_spaces)
                just_fids = filter(lambda a: a != 'XXX', just_fids)
                fids_for_cube += just_fids
            cube = pt.data_cube[fids_for_cube]
            periods = pt.desired_periods[::-1]
            cube = cube.loc[:, :, periods]
            output_filename = "%s.csv" % bus_type
            for key in cube.minor_axis:
                df = cube.minor_xs(key)
                df.to_csv(output_filename, mode='a', encoding='utf-8')
        return

    def get_groups_owning_cusips(self, cusip_list):
        bt_grp_dict = {}
        for tag, bt in self.business_types.items():
            if bt:
                cusip_dict = grp_cusips = bt.get_groups_owning_cusips(cusip_list)
                bt_grp_dict[tag] = cusip_dict
        return bt_grp_dict


    def get_companies(self, tag, y_or_q):
        if self.business_types[tag] is not None:
            if y_or_q in self.business_types[tag].period_types:
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
