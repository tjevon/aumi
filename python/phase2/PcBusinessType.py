from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')


class PcBusinessType(BusinessType):
    """Class containing all PC companies and their data"""
    def __init__(self, bt_tag, period_idx, common_tag_list, bt_tag_list):
        super(PcBusinessType, self).__init__()
        self.bt_tag = bt_tag
        self.period_idx = period_idx
        self.common_tag_list = common_tag_list
        self.bt_tag_list = bt_tag_list
        self.complete_tag_list = self.common_tag_list + self.bt_tag_list
        return

#    def convert_csvs_to_raw_df(self, data_dir, file_names):
#        remaining_file_names = self.convert_common_csvs_to_raw_df(data_dir, file_names, self.complete_tag_list)
#        return

    def construct_data_cube(self, data_dir, file_names, template_wb):
#        self.convert_csvs_to_raw_df(data_dir, file_names)
        self.convert_common_csvs_to_raw_df(data_dir, file_names, self.complete_tag_list)
        self.construct_my_data_cube(template_wb, self.complete_tag_list, self.period_idx)
        return

    def get_bt_tag(self):
        return self.bt_tag
