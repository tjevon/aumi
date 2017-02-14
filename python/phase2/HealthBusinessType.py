from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')


class HealthBusinessType(BusinessType):
    """Class containing all Health companies and their data"""
    def __init__(self):
        logger.info("Enter")
        super(HealthBusinessType, self).__init__()
        self.raw_df = None
        self.data_cube = None
        logger.info("Leave")
        return

    def convert_csvs_to_raw_df(self, data_dir, file_names):
        logger.info("Enter")
        remaining_file_names = super(HealthBusinessType, self).convert_csvs_to_raw_df(data_dir, file_names)
        self.handle_remaining_files(data_dir, remaining_file_names)
        logger.info("Leave")
        return

    def handle_remaining_files(self, data_dir, file_names):
        logger.info("Enter")
        for file_name in file_names:
            if any(s in file_name for s in HEALTH_TEMPLATE_TAGS):
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

    def construct_data_cube(self, template_wb):
        self.construct_my_data_cube(template_wb, COMMON_TEMPLATE_TAGS + HEALTH_TEMPLATE_TAGS)
        return

    def get_bt_tag(self):
        return HEALTH_tag
