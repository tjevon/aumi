from __future__ import print_function

from MasterProspectList import *
from TearSheetGenerator import *
from TemplateWorkBook import *

from infra.cmdline_logging import parse, setup_default_options, setup_logging

data_dir = os.getenv('DATAPATH', './')
log_dir = os.getenv('LOGPATH', './')

def construct_mpl(template_obj, target_obj, data_dir):
    logger.debug("Enter")
    rv_mpl = MasterProspectList(template_obj, target_obj, data_dir)
    logger.debug("Leave")
    return rv_mpl

def construct_tearsheet_generator(template_obj):
    logger.debug("Enter")
    rv_gen = TearSheetGenerator(data_dir, template_obj)
    logger.debug("Leave")
    return rv_gen

def construct_template_object(data_dir):
    logger.debug("Enter")
    rv_template_obj = TemplateWorkBook(data_dir)
    logger.debug("Leave")
    return rv_template_obj

# Called by main
################################################################################
if __name__ == '__main__':
    logger.debug("Enter")
    parser = setup_default_options()
    options = parse(parser)
    logger = setup_logging("AMB_mpl.log", options.logger_level)

    template_obj = construct_template_object(data_dir)
    target_obj = construct_tearsheet_generator(template_obj)
    mpl = construct_mpl(template_obj, target_obj, data_dir)

    company_dict = {}

#    mpl.build_trio_scorecard()
#    mpl.build_specialty_cube()
    company_dict[PC_tag] = mpl.get_candidates(PC_tag)
    company_dict[LIFE_tag] = mpl.get_candidates(LIFE_tag)
    company_dict[HEALTH_tag] = mpl.get_candidates(HEALTH_tag)

    ### build tearsheets for top candidates
    target_obj.build_tearsheets(company_dict, mpl)

    logger.debug("Leave")

