from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')


class LifeBusinessType(BusinessType):
    """Class containing all Life companies and their data"""
    def __init__(self):
        logger.info("Enter")
        super(LifeBusinessType, self).__init__()
        logger.info("Leave")
        return

    def convert_csvs_to_raw_df(self, data_dir, file_names, common_tag_list, life_tag_list):
        logger.info("Enter")
        remaining_file_names = self.convert_common_csvs_to_raw_df(data_dir, file_names, common_tag_list)
        self.handle_remaining_csvs(data_dir, remaining_file_names, life_tag_list)
        logger.info("Leave")
        return

    def handle_remaining_csvs(self, data_dir, file_names, life_tag_list):
        logger.info("Enter")
        for file_name in file_names:
            if any(s in file_name for s in life_tag_list):
                csv_filename = data_dir + "\\" + file_name
                df = self.load_df(csv_filename)
                # TODO: should we check to make sure fids don't already exist?
                try:
                    self.raw_df = pd.concat([self.raw_df, df], axis=1)
                except NameError:
                    self.raw_df = df
            else:
                logger.info("filename %s not part of COMMON or LIFE", file_name)
        logger.info("Leave")
        return

    def construct_data_cube(self, data_dir, file_names, template_wb):
        self.convert_csvs_to_raw_df(data_dir, file_names, COMMON_TEMPLATE_TAGS, LIFE_TEMPLATE_TAGS)
        self.construct_my_data_cube(template_wb, COMMON_TEMPLATE_TAGS + LIFE_TEMPLATE_TAGS, YEARLY_FID_IDX)
        return

    def get_bt_tag(self):
        return LIFE_tag
