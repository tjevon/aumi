from __future__ import print_function

from MasterProspectList import *
from TearSheetGenerator import *
from TemplateWorkBook import *

from infra.cmdline_logging import parse, setup_default_options, setup_logging

log_dir = os.getenv('LOGPATH', './')


def construct_mpl():
    logger.debug("Enter")
    rv_mpl = MasterProspectList(template_obj, target_obj)
    logger.debug("Leave")
    return rv_mpl


def construct_tearsheet_generator():
    logger.debug("Enter")
    rv_gen = TearSheetGenerator(template_obj)
    logger.debug("Leave")
    return rv_gen


def construct_template_object():
    logger.debug("Enter")
    rv_template_obj = TemplateWorkBook()
    logger.debug("Leave")
    return rv_template_obj

# Called by main
################################################################################
if __name__ == '__main__':
    logger.debug("Enter")
    parser = setup_default_options()
    options = parse(parser)
    logger = setup_logging("AMB_mpl.log", options.logger_level)

    template_obj = construct_template_object()
    target_obj = construct_tearsheet_generator()
    mpl = construct_mpl()
    mpl.build_projections_cube()
    mpl.build_mpl_sheet()

    comp_yearly_dict = {}
    comp_qtrly_dict = dict()

#    mpl.build_trio_scorecard()
#    mpl.build_specialty_cube()
#    comp_yearly_dict {PC_tag:dict}
    comp_yearly_dict[PC_tag] = mpl.get_candidates(PC_tag, YEARLY_IDX)
    comp_yearly_dict[LIFE_tag] = mpl.get_candidates(LIFE_tag, YEARLY_IDX)
    comp_yearly_dict[HEALTH_tag] = mpl.get_candidates(HEALTH_tag, YEARLY_IDX)

    target_obj.build_tearsheets(comp_yearly_dict, mpl, YEARLY_IDX)

#    comp_qtrly_dict[PC_tag] = mpl.get_candidates(PC_tag, QUARTERLY_IDX)
#    comp_qtrly_dict[LIFE_tag] = mpl.get_candidates(LIFE_tag, QUARTERLY_IDX)
#    comp_qtrly_dict[HEALTH_tag] = mpl.get_candidates(HEALTH_tag, QUARTERLY_IDX)

    target_obj.build_tearsheets(comp_qtrly_dict, mpl, QUARTERLY_IDX)

    target_obj.target_wb.app.calculate()
    target_obj.target_wb.app.calculation = 'automatic'
    logger.debug("Leave")
