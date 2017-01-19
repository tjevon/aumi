from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')

class LifeBusinessType(BusinessType):
    """Class containing all Life companies and their data"""
    def __init__(self):
        logger.info("Enter")
        super(LifeBusinessType, self).__init__()
        logger.info("Leave")
        pass

    def process_csvs(self, data_dir, file_names):
        logger.info("Enter")
        remaining_file_names = super(LifeBusinessType, self).process_csvs(data_dir, file_names)
        self.handle_remaining_files(data_dir, remaining_file_names)
        logger.info("Leave")

    def handle_remaining_files(self, data_dir, file_names):
        logger.info("Enter")
        logger.info("Leave")
        pass
