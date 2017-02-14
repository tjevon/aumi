from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')


class PcBusinessType(BusinessType):
    """Class containing all PC companies and their data"""
    def __init__(self):
        super(PcBusinessType, self).__init__()
        self.raw_df = None
        self.data_cube = None
        return

    def convert_csvs_to_raw_df(self, data_dir, file_names):
        remaining_file_names = super(PcBusinessType, self).convert_csvs_to_raw_df(data_dir, file_names)
        self.handle_remaining_csvs(data_dir, remaining_file_names)
        return

    def handle_remaining_csvs(self, data_dir, file_names):
        for file_name in file_names:
            if any(s in file_name for s in PC_TEMPLATE_TAGS):
                csv_filename = data_dir + "\\" + file_name
                df = self.load_df(csv_filename)
                # TODO: should we check to make sure fids don't already exist?
                try:
                    self.raw_df = pd.concat([self.raw_df, df], axis=1)
                except NameError:
                    self.raw_df = df
            else:
                logger.info("Filename %s not part of COMMON or PC", file_name)
        logger.info("Leave")
        return

    def construct_data_cube(self, template_wb):
        self.construct_my_data_cube(template_wb, COMMON_TEMPLATE_TAGS + PC_TEMPLATE_TAGS)
        return

    def get_bt_tag(self):
        return PC_tag
