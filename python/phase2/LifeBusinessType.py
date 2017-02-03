from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')

class LifeBusinessType(BusinessType):
    """Class containing all Life companies and their data"""
    def __init__(self):
        logger.info("Enter")
        super(LifeBusinessType, self).__init__()
        self.LIFE_raw_df = None
        logger.info("Leave")
        return

    def process_csvs(self, data_dir, file_names):
        logger.info("Enter")
        remaining_file_names = super(LifeBusinessType, self).process_csvs(data_dir, file_names)
        self.handle_remaining_csvs(data_dir, remaining_file_names)
        logger.info("Leave")
        return

    def handle_remaining_csvs(self, data_dir, file_names):
        logger.info("Enter")
        for file_name in file_names:
            if any(s in file_name for s in LIFE_TEMPLATE_TAGS):
                csv_filename = data_dir + "\\" + file_name
                the_df = self.load_df(csv_filename)
                # TODO: should we check to make sure fids don't already exist?
                try:
                    self.LIFE_raw_df = pd.concat([self.LIFE_raw_df, the_df], axis=1)
                except NameError:
                    self.LIFE_raw_df = the_df
            else:
                logger.info("filename %s not part of COMMON or LIFE", file_name)
        logger.info("Leave")
        return


    def construct_data_cubes(self, template_wb):
        self.construct_select_data_cubes(template_wb, COMMON_TEMPLATE_TAGS)
        self.construct_select_data_cubes(template_wb, LIFE_TEMPLATE_TAGS)
        return

    def get_bt_tag(self):
        return LIFE_tag

    def get_derived_df(self, sheet):
        valid = True
        the_df = self.LIFE_raw_df
        return (valid, the_df)

