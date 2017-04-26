from __future__ import print_function
import numpy as np
from PeriodType import *
from collections import defaultdict

logger = logging.getLogger('twolane')


class BusinessType(object):
    """Base class for PC, Life, Health"""

    def __init__(self, bt_tag, template_obj):
        self.period_types = {YEARLY_IDX: None, QUARTERLY_IDX: None}
        self.bt_tag = bt_tag
        self.company_to_grp = {}
        self.group_to_company = defaultdict(list)
        self.company_info_df = None

        self.build_company_to_grp()

        self.period_types[YEARLY_IDX] = PeriodType(self, YEARLY_IDX, template_obj)
        self.period_types[QUARTERLY_IDX] = PeriodType(self, QUARTERLY_IDX, template_obj)

        return

    def get_group_names(self, group_numbers):
        df = self.company_info_df.ix[group_numbers]
        group_names = df['CO00231'].tolist()
        return group_names

    def get_company_info(self, y_or_q):
        amb_numbers = list(self.period_types[y_or_q].grp_unaf)
        df = self.company_info_df.ix[amb_numbers]
        group_info_list = df['CO00231'].tolist()
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
        amb_number = 'CO00002'
        group_fid = 'CO00023'
        parent_fid = 'CO00169'
        ultimate_fid = 'CO00170'
        csv_filename = csv_data_dir + "\\" + file_list[0]
        df_map = pd.read_csv(csv_filename, header=2, index_col=0, dtype='unicode', thousands=",")
        df_map = df_map[[amb_number, group_fid, parent_fid, ultimate_fid]].drop(['AMB#'])
        df_map[group_fid] = np.where(df_map[group_fid] == '000000', df_map[amb_number], df_map[group_fid])

        df_map[amb_number] = df_map[amb_number].apply(lambda x: x.zfill(6))
        df_map[group_fid] = df_map[group_fid].apply(lambda x: x.zfill(6))
        df_map[parent_fid] = df_map[parent_fid].apply(lambda x: x.zfill(6))
        df_map[ultimate_fid] = df_map[ultimate_fid].apply(lambda x: x.zfill(6))

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
        self.company_info_df[amb_number] = self.company_info_df[amb_number].apply(lambda x: x.zfill(6))
        self.company_info_df[group_fid] = self.company_info_df[group_fid].apply(lambda x: x.zfill(6))
        self.company_info_df[parent_fid] = self.company_info_df[parent_fid].apply(lambda x: x.zfill(6))
        self.company_info_df[ultimate_fid] = self.company_info_df[ultimate_fid].apply(lambda x: x.zfill(6))

        df_map = df_map.loc[df_map[group_fid].isin(self.company_info_df.index)]

#        my_tuple_df = df[[amb_number, group_fid]].drop(['AMB#'])
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
