from __future__ import print_function

from MasterProspectList import *
from TearSheetGenerator import *

from infra.cmdline_logging import parse, setup_default_options, setup_logging

data_dir = os.getenv('DATAPATH', './')
log_dir = os.getenv('LOGPATH', './')

def construct_mpl():
    logger.debug("Enter")
    the_mpl = MasterProspectList()
    the_mpl.get_data_from_files(data_dir)
    logger.debug("Leave")
    return the_mpl

def construct_tearsheet_generator():
    logger.debug("Enter")
    the_gen = TearSheetGenerator(data_dir)
    logger.debug("Leave")
    return the_gen

def get_arbitrary_subset(companies):
    tmp_set = set()
    for i in range(0,5):
        tmp_set.add(companies.pop())
    return tmp_set

# Called by main
################################################################################
if __name__ == '__main__':
    logger.debug("Enter")
    parser = setup_default_options()
    options = parse(parser)
    logger = setup_logging("AMB_mpl.log", options.logger_level)
    my_mpl = construct_mpl()
    my_gen = construct_tearsheet_generator()

    # for testing tearsheet creation only,,, create small set of companies to print
    tmp_set = get_arbitrary_subset(my_mpl.get_pc_companies())

    my_gen.build_pc_tearsheets(tmp_set, my_mpl)

    logger.debug("Leave")

