from __future__ import print_function

from BusinessType import *

logger = logging.getLogger('twolane')

class PcBusinessType(BusinessType):
    """Class containing all PC companies and their data"""
    def __init__(self):
        super(PcBusinessType, self).__init__()
        self.SoI_raw_df = None
        pass

    def process_csvs(self, data_dir, file_names):
        remaining_file_names = super(PcBusinessType, self).process_csvs(data_dir, file_names)
        self.handle_remaining_csvs(data_dir, remaining_file_names)

    def handle_remaining_csvs(self,data_dir, file_names):
        for file_name in file_names:
            csv_filename = data_dir + "\\" + file_name
            if file_name.find(SoI_files) != -1:
                self.SoI_raw_df = self.load_df(csv_filename)
        logger.info("Leave")
        pass

    def construct_data_cubes(self, template_wb):
        remaining_file_names = super(PcBusinessType, self).construct_data_cubes(template_wb, COMMON_TEMPLATE_SHEETS)
        remaining_file_names = super(PcBusinessType, self).construct_data_cubes(template_wb, PC_TEMPLATE_SHEETS)
        pass

    def get_template_column(self, tag):
        # TODO: clean this up
        if tag == "E07":
            return 'C7:C54'
        elif tag == "SI01":
            return 'B5:B67'
        elif tag == "Assets":
            return 'B2:B58'
        elif tag == "SoI":
            return 'B3:B70'
        elif tag == "CashFlow":
            return 'B6:B50'
        else:
            logger.error("Looking for unknown tag: %s", tag)
            return 'A1'

    def get_derived_df(self, sheet):
        the_df = None
        valid = True
        if sheet == "SoI":
            the_df = self.SoI_raw_df
        else:
            valid = False
            logger.error("No sheet available in base class: %s", sheet)
        return (valid, the_df)

    def set_derived_cube(self, sheet, the_cube):
        rv = True
        if sheet == "SoI":
            self.SoI_cube = the_cube
        else:
            logger.error("No cube for: %s", sheet)
            rv = False
        return rv

