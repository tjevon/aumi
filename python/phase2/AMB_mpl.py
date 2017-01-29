from __future__ import print_function

from MasterProspectList import *
from TearSheetGenerator import *
from TemplateWorkBook import *

from infra.cmdline_logging import parse, setup_default_options, setup_logging

data_dir = os.getenv('DATAPATH', './')
log_dir = os.getenv('LOGPATH', './')

def construct_mpl(template_obj):
    logger.debug("Enter")
    the_mpl = MasterProspectList(template_obj)
    the_mpl.get_data_from_files(data_dir)
    logger.debug("Leave")
    return the_mpl

def construct_tearsheet_generator(template_obj):
    logger.debug("Enter")
    the_gen = TearSheetGenerator(data_dir, template_obj)
    logger.debug("Leave")
    return the_gen

def get_arbitrary_subset(companies):
    tmp_set = set()
#    for i in range(0,len(companies)):
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

    my_template_obj = TemplateWorkBook(data_dir)
    my_mpl = construct_mpl(my_template_obj)

    # for testing tearsheet creation only,,, create small set of companies to print
    tmp_set = get_arbitrary_subset(my_mpl.get_pc_companies())

    my_gen = construct_tearsheet_generator(my_template_obj)
    my_gen.build_tearsheets(tmp_set, my_mpl)

    logger.debug("Leave")

