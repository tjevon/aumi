from __future__ import print_function
import numpy as np
from PeriodType import *
from collections import defaultdict

logger = logging.getLogger('twolane')


class BusinessType(object):
    """Base class for PC, Life, Health"""

    def __init__(self, bt_tag, template_obj, quarterly_arg):
        self.period_types = {YEARLY_IDX: None, QUARTERLY_IDX: None}
        self.bt_tag = bt_tag
        self.company_to_grp = {}
        self.group_to_company = defaultdict(list)
        self.company_info_df = None

        self.build_company_to_grp()

        self.period_types[QUARTERLY_IDX] = PeriodType(self, QUARTERLY_IDX, template_obj,
                                                      quarterly_arg)
        self.period_types[YEARLY_IDX] = PeriodType(self, YEARLY_IDX, template_obj,
                                                   quarterly_arg)
        if quarterly_arg != 'Annual':
            self.compute_YTD(template_obj)
        return

    def compute_YTD(self,template_obj):
        yearly_fid_list = template_obj.get_display_fid_list(Liquid_Assets_tag, self.bt_tag, YEARLY_IDX)

        yearly_fid_list = template_obj.get_display_fid_list(Liquid_Assets_tag, self.bt_tag, YEARLY_IDX)
        yearly_slice = self.period_types[YEARLY_IDX].data_cube.minor_xs(self.period_types[YEARLY_IDX].desired_periods[0])
        yearly_liquid_assets_df = yearly_slice[yearly_fid_list]

        qtrly_fid_list = template_obj.get_display_fid_list(Liquid_Assets_tag, self.bt_tag, QUARTERLY_IDX)
        qtrly_slice = self.period_types[QUARTERLY_IDX].data_cube.minor_xs(self.period_types[QUARTERLY_IDX].desired_periods[0])
        qtrly_liquid_assets_df = qtrly_slice[qtrly_fid_list]

        yearly_liquid_assets_df.columns = qtrly_liquid_assets_df.columns

        pct_label = self.period_types
        qtrly_label = self.period_types[QUARTERLY_IDX].desired_periods[0]
        yearly_label = self.period_types[YEARLY_IDX].desired_periods[0]
        pct_label = qtrly_label + '.' + yearly_label
        ordered_dates = [qtrly_label, yearly_label, pct_label]

        df_dict = {}
        df_dict[ordered_dates[0]] = qtrly_liquid_assets_df
        df_dict[ordered_dates[1]] = yearly_liquid_assets_df

        tmp_cube = pd.Panel(df_dict)
        tmp_cube = tmp_cube.reindex(ordered_dates[0:2])
        df_dict = {}
        for co in tmp_cube.major_axis:
            tmp_df = tmp_cube.major_xs(co)
            pct_chg = tmp_df[ordered_dates[0]]/tmp_df[ordered_dates[1]]
            tmp_df[pct_label] = pct_chg
            df_dict[co] = tmp_df

        tmp_cube = pd.Panel(df_dict)
        tmp_cube = tmp_cube.swapaxes(0,1)
        self.period_types[QUARTERLY_IDX].ytd_cube = tmp_cube
        return

    def get_groups_owning_cusips(self, cusip_list):
        cusip_dict = {}
        for y_or_q, pt in self.period_types.items():
            for yr, pos_tup in pt.position_cube.items():
                bonds_owned = pos_tup[0]
                for grp, grp_df in bonds_owned.items():
                    df = grp_df.loc[grp_df['cusip'].isin(cusip_list)]
                    if df.size != 0:
                        cusip_dict[grp] = df
                stocks_owned = pos_tup[1]
                for grp, grp_df in stocks_owned.items():
                    df = grp_df.loc[grp_df['cusip'].isin(cusip_list)]
                    if df.size != 0:
                        cusip_dict[grp] = df
                        logger.error("grp: %s located %d positions", grp, df.shape[0])
        return cusip_dict

    def get_group_info_df(self, group_numbers):
        df = self.company_info_df.loc[group_numbers, :]
        group_info_df = df.loc[:, [COMPANY_NAME_FID, COMPANY_CITY_FID, COMPANY_STATE_FID]]
        return group_info_df

    def get_group_names(self, group_numbers):
        df = self.company_info_df.loc[group_numbers, :]
        test = df.loc[:, [COMPANY_NAME_FID, COMPANY_CITY_FID, COMPANY_STATE_FID]]
        group_names = df[COMPANY_NAME_FID].tolist()
        return group_names

    def get_company_info(self, y_or_q):
        amb_numbers = list(self.period_types[y_or_q].grp_unaf)
        df = self.company_info_df.loc[amb_numbers, :]
        group_info_list = df[COMPANY_NAME_FID].tolist()
        group_info_dict = dict(zip(amb_numbers, group_info_list))
        return group_info_dict

    def build_company_to_grp(self):
        csv_data_dir = DATA_DIR + COMPANY_MAP_DIR
        file_dict = PeriodType.get_file_dict(csv_data_dir)
        if self.bt_tag not in file_dict:
            return

        file_list = file_dict[self.bt_tag]
        if len(file_list) > 1:
            logger.error("file_list length > 1")
            return
        csv_data_dir = DATA_DIR + COMPANY_MAP_DIR
        csv_filename = csv_data_dir + "\\" + file_list[0]
        df_map = pd.read_csv(csv_filename, header=2, index_col=0, dtype='unicode', thousands=",")
        df_map = df_map[[AMB_NUMBER, GROUP_FID, PARENT_FID, ULTIMATE_FID]].drop(['AMB#'])
        df_map[GROUP_FID] = np.where(df_map[GROUP_FID] == '000000', df_map[AMB_NUMBER], df_map[GROUP_FID])

        df_map[AMB_NUMBER] = df_map[AMB_NUMBER].apply(lambda x: x.zfill(6))
        df_map[GROUP_FID] = df_map[GROUP_FID].apply(lambda x: x.zfill(6))
        df_map[PARENT_FID] = df_map[PARENT_FID].apply(lambda x: x.zfill(6))
        df_map[ULTIMATE_FID] = df_map[ULTIMATE_FID].apply(lambda x: x.zfill(6))

        csv_data_dir = DATA_DIR + COMPANY_INFO_DIR
        file_dict = PeriodType.get_file_dict(csv_data_dir)
        if self.bt_tag not in file_dict:
            return
        file_list = file_dict[self.bt_tag]
        csv_filename = csv_data_dir + "\\" + file_list[0]
        self.company_info_df = pd.read_csv(csv_filename, header=2, index_col=0, dtype='unicode', thousands=",")
        self.company_info_df = self.company_info_df.drop(['AMB#'])
        new_index = self.company_info_df.index
        new_index = [str(x).zfill(6) for x in new_index]
        self.company_info_df = self.company_info_df.set_index([new_index])
        self.company_info_df[AMB_NUMBER] = self.company_info_df[AMB_NUMBER].apply(lambda x: x.zfill(6))
        self.company_info_df[GROUP_FID] = self.company_info_df[GROUP_FID].apply(lambda x: x.zfill(6))
        self.company_info_df[PARENT_FID] = self.company_info_df[PARENT_FID].apply(lambda x: x.zfill(6))
        self.company_info_df[ULTIMATE_FID] = self.company_info_df[ULTIMATE_FID].apply(lambda x: x.zfill(6))

        df_map = df_map.loc[df_map[GROUP_FID].isin(self.company_info_df.index)]

#        my_tuple_df = df[[AMB_NUMBER, GROUP_FID]].drop(['AMB#'])
        for idx, row in df_map.iterrows():
            self.company_to_grp[row[0]] = row[1]
        for k, v in self.company_to_grp.items():
            if v == '000000':
                self.group_to_company[k].append(k)
            else:
                self.group_to_company[v].append(k)

#        self.period_types[QUARTERLY_IDX].set_group_to_company(self.group_to_company)

        return

    def get_bt_tag(self):
        return self.bt_tag

# def get_specialty_cube(self, fids):
#     cube = self.data_cube[fids]
#     return cube
