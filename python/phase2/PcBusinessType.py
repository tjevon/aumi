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
        self.handle_remaining_csvs(data_dir, remaining_file_names)

    def handle_remaining_csvs(self,data_dir, file_names):
        for file_name in file_names:
            csv_filename = data_dir + "\\" + file_name
            if file_name.find(SoI_files) != -1:
                self.soi_df = self.load_df(csv_filename)
#            handle_remaining_cubes()
        logger.info("Leave")
        pass

    def construct_data_cubes(self, template_wb):
        remaining_file_names = super(PcBusinessType, self).construct_data_cubes(template_wb)
#        self.handle_remaining_cubes(template_wb, remaining_file_names)

    def get_template_column(self, tag):
        if tag == "E07":
            return 'C7:C54'
        elif tag == "SI01":
            return 'B5:B67'
        return 'A1'


    def handle_remaining_cubes(self,template_wb, file_names):
        for file_name in file_names:
            pass
        logger.info("Leave")
        pass

