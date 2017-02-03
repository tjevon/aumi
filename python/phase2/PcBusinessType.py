from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')

class PcBusinessType(BusinessType):
    """Class containing all PC companies and their data"""
    def __init__(self):
        super(PcBusinessType, self).__init__()
        self.PC_raw_df = None
        return

    def process_csvs(self, data_dir, file_names):
        remaining_file_names = super(PcBusinessType, self).process_csvs(data_dir, file_names)
        self.handle_remaining_csvs(data_dir, remaining_file_names)
        return

    def handle_remaining_csvs(self,data_dir, file_names):
        for file_name in file_names:
            if any(s in file_name for s in PC_TEMPLATE_TAGS):
                csv_filename = data_dir + "\\" + file_name
                the_df = self.load_df(csv_filename)
                # TODO: should we check to make sure fids don't already exist?
                try:
                    self.PC_raw_df = pd.concat([self.PC_raw_df, the_df], axis=1)
                except NameError:
                    self.PC_raw_df = the_df
            else:
                logger.info("Filename %s not part of COMMON or PC", file_name)
        logger.info("Leave")
        return

    def construct_data_cubes(self, template_wb):
        remaining_file_names = self.construct_select_data_cubes(template_wb, COMMON_TEMPLATE_TAGS)
        remaining_file_names = self.construct_select_data_cubes(template_wb, PC_TEMPLATE_TAGS)
        return

    def get_bt_tag(self):
        return PC_tag

    def get_derived_df(self, sheet):
        the_df = None
        valid = True
        the_df = self.PC_raw_df
        return (valid, the_df)
