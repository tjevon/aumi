from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')

class PcBusinessType(BusinessType):
    """Class containing all PC companies and their data"""
    def __init__(self):
        super(PcBusinessType, self).__init__()
        self.soi_df = None
        pass

    def process_csvs(self, data_dir, file_names):
        remaining_file_names = super(PcBusinessType, self).process_csvs(data_dir, file_names)
        self.handle_remaining_files(data_dir, remaining_file_names)

    def handle_remaining_files(self,data_dir, file_names):
        for file_name in file_names:
            csv_filename = data_dir + "\\" + file_name
            if file_name.find(soi) != -1:
                self.soi_df = self.load_df(csv_filename)
        logger.info("Leave")
        pass

