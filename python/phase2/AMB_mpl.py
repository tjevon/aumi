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


def construct_tearsheet_generator(template_obj):
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
    parser.add_option("-q", "--quarterly", action="store", type="string", dest="quarterly",
                      default="None",
                      help="specify quarter for quarterly report",
                      metavar="default value of None, format (YYYY-Q)")
    options = parse(parser)
    quarterly = options.quarterly
    logger = setup_logging("AMB_mpl.log", options.logger_level)

    template_obj = construct_template_object()
    target_obj = construct_tearsheet_generator(template_obj)
    mpl = MasterProspectList(template_obj, target_obj, quarterly)
    asset_mgr = "StoneHarbor"
#    target_obj.build_cusip_ownership_sheet(asset_mgr, mpl)
    mpl.build_projections_cube()
    mpl.build_mpl_sheet('Liq MPL', Liquid_Assets_tag)
    mpl.build_mpl_sheet('Illiq MPL', E07_tag)
    mpl.build_industry_info_sheet('PC_Info', Asset_Alloc_tag, PC_tag)
    mpl.build_industry_info_sheet('LIFE Info', Asset_Alloc_tag, LIFE_tag)
    mpl.build_industry_info_sheet('HEALTH Info', Asset_Alloc_tag, HEALTH_tag)

    comp_yearly_dict = {}
    comp_qtrly_dict = dict()

    debug = False

#    comp_yearly_dict[PC_tag] = mpl.get_candidates(PC_tag, YEARLY_IDX)
#    comp_yearly_dict[LIFE_tag] = mpl.get_candidates(LIFE_tag, YEARLY_IDX)
#    comp_yearly_dict[HEALTH_tag] = mpl.get_candidates(HEALTH_tag, YEARLY_IDX)
    target_obj.build_tearsheets(comp_yearly_dict, mpl, YEARLY_IDX, debug)

    comp_qtrly_dict[PC_tag] = mpl.get_candidates(PC_tag, QUARTERLY_IDX)
#    comp_qtrly_dict[LIFE_tag] = mpl.get_candidates(LIFE_tag, QUARTERLY_IDX)
#    comp_qtrly_dict[HEALTH_tag] = mpl.get_candidates(HEALTH_tag, QUARTERLY_IDX)
    target_obj.build_tearsheets(comp_qtrly_dict, mpl, QUARTERLY_IDX, debug)

    target_obj.target_wb.app.calculate()
    target_obj.target_wb.app.calculation = 'automatic'
    logger.debug("Leave")
