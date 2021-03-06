from __future__ import print_function

import xlwings as xw

from AMB_defines import *

logger = logging.getLogger('twolane')


class TemplateWorkBook(object):

    template_section_ranges = {
        SI01_tag: ('A', 3, 154),
        E07_tag: ('A',  3, 50),
        BA_Acq_tag: ('A',  3, 50),
        BA_Disp_tag: ('A',  3, 50),
        E10_tag: ('A',  3, 59),
        Assets_tag: ('A', 3, 87),
        CashFlow_tag: ('A', 3, 48),
        SI05_07_tag: ('A', 3, 106),
        Liquid_Assets_tag: ('A', 3, 60),
        Liq_Acq_tag: ('A', 3, 60),
        Liq_Disp_tag: ('A', 3, 60),
        Asset_Alloc_tag: ('A', 3, 78),
        SoI_tag: ('A', 3, 69),
        SoO_tag: ('A', 3, 70),
        SoR_tag: ('A', 3, 62),
        IRIS1_tag: ('A', 3, 30),
        IRIS2_tag: ('A', 3, 28),
        Liab1_tag: ('A', 3, 52),
        Liab2_tag: ('A', 3, 67),
        Liab3_tag: ('A', 3, 41),
        Real_Estate_tag: ('A', 3, 7),
        CR_tag: ('A', 3, 4)
    }

    # fid located in first column, display filter in second
    template_BT_cols = {
        PC_tag: ('B', 'C', 'M'),
        LIFE_tag: ('D', 'E', 'N'),
        HEALTH_tag: ('F', 'G', 'O')
    }

    SECTION_TARGET_COLUMN_LOCATION = "J"
    SECTION_TARGET_DATA_LOCATION = "K"
    SECTION_TARGET_ROW_LOCATION = "L"

    QTRLY_PROJECTION_TYPE_COL = 'P'

    def __init__(self):
        self.xl_wb = self.get_template_workbook()

    @staticmethod
    def get_template_workbook():
        template_data_dir = XL_DIR + TEMPLATE_DIR
        workbook_file = template_data_dir + "\\" + TEMPLATE_FILENAME
        template_wb = xw.Book(workbook_file)
        available_sheets = []
        for i in template_wb.sheets:
            logger.debug(i.name)
            available_sheets.append(i.name)
        logger.debug("Have Sheets:")

        for sheet_name in COMMON_TEMPLATE_TAGS + PC_TEMPLATE_TAGS + LIFE_TEMPLATE_TAGS + HEALTH_TEMPLATE_TAGS:
            if sheet_name not in available_sheets:
                logger.fatal("Template %s does not contain %s", workbook_file, sheet_name)
                exit()
        return template_wb

    def get_template_sheet(self, tag):
        rv_template_sheet = self.xl_wb.sheets(tag)
        return rv_template_sheet

    def get_formula(self, sheet, cell):
        template_sheet = self.get_template_sheet(sheet)
        formula = template_sheet.range(cell).formula
        return formula

    def get_full_fid_list(self, tag, bt_tag, y_or_q=YEARLY_IDX):
        template_sheet = self.get_template_sheet(tag)
        range_info = self.template_section_ranges[tag]

        a_cols = []
        for i in range(range_info[1], range_info[2]+1):
            a_cols.append(str(range_info[0]) + str(i))

        fid_col = self.template_BT_cols[bt_tag][y_or_q]
        cells = fid_col + str(range_info[1]+1) + ':' + fid_col + str(range_info[2])
        fid_list = template_sheet.range(cells).options(transpose=True).value

        fid_dict = {}
        if type(fid_list) is list:
            fid_dict = dict(zip(a_cols[1:], fid_list))
            return fid_list, fid_dict
        else:
            tmp_fid_list = list()
            tmp_fid_list.append(fid_list)
            fid_dict[a_cols[1]] = fid_list
            return tmp_fid_list, fid_dict

    def get_display_fid_list(self, tag, bt_tag, y_or_q, which_filter=DISPLAY_TS_SECTION):
        fid_list = self.get_full_fid_list(tag, bt_tag, y_or_q)[0]
        rv = self.filter_list(tag, bt_tag, fid_list, 1, which_filter)[0]
        return rv

    def get_projection_info(self, tag, bt_tag, y_or_q, which_filter, debug=False):
        fid_list = self.get_full_fid_list(tag, bt_tag, y_or_q)[0]
        if debug:
            rv = fid_list, {}
        else:
            rv = self.filter_list(tag, bt_tag, fid_list, 1, which_filter)
        return rv

    def get_row_labels(self, tag, bt_tag, display_type, debug_row=0, debug=False):
        template_sheet = self.get_template_sheet(tag)
        range_info = self.template_section_ranges[tag]

        cells = str(range_info[0]) + str(range_info[1]) + ':' + str(range_info[0]) + str(range_info[2])
        row_labels = template_sheet.range(cells).options(ndim=2).value


        target_location_row = 0
        if display_type == DISPLAY_TS_PROJ:
            target_location_row = range_info[1] + 1
        else:
            target_location_row = range_info[1]
        row = 0
        if debug == False:
            row_labels = self.filter_list(tag, bt_tag, row_labels, 0, display_type)[0]
            cells = self.SECTION_TARGET_ROW_LOCATION + str(target_location_row)
            row = template_sheet.range(cells).value
        else:
            row = debug_row

        cells = self.SECTION_TARGET_COLUMN_LOCATION + str(target_location_row)
        row_label_column = template_sheet.range(cells).value
        cells = self.SECTION_TARGET_DATA_LOCATION + str(target_location_row)
        data_column = template_sheet.range(cells).value

        return row_labels, row_label_column, data_column, row

    def filter_list(self, tag, bt_tag, to_filter, offset, which_filter=DO_NOT_DISPLAY):
        template_sheet = self.get_template_sheet(tag)
        range_info = self.template_section_ranges[tag]
        display_filter_col = self.template_BT_cols[bt_tag][DISPLAY_IDX]
        cells = display_filter_col + str(range_info[1] + offset) + ':' + display_filter_col + str(range_info[2])
        filters = template_sheet.range(cells).options(ndim=2).value
        filters = [item for sublist in filters for item in sublist]
        filters = [0 if x is None else x for x in filters]
        filters = map(int, filters)
        rv = []

        # if which_filter == DISPLAY_TS_PROJ or which_filter == DISPLAY_MPL:
        if which_filter == DISPLAY_MPL:
            proj_type_col = self.QTRLY_PROJECTION_TYPE_COL
            proj_cat_range = proj_type_col + str(range_info[1]+offset) + ':' + proj_type_col + str(range_info[2])
            cats = template_sheet.range(proj_cat_range).options(ndim=2).value

        proj_dict = {}
        for filter_cell, x in zip(filters, range(len(filters))):
            if which_filter == DISPLAY_TS_SECTION:
                if filter_cell & DISPLAY_TS_SECTION:
                    rv.append(to_filter[x])
            elif which_filter == DISPLAY_TS_PROJ:
                if filter_cell & DISPLAY_TS_PROJ:
#                    proj_dict[to_filter[x]] = cats[x][0]
                    rv.append(to_filter[x])
            elif which_filter == DISPLAY_MPL:
                if filter_cell & DISPLAY_MPL:
                    proj_dict[to_filter[x]] = cats[x][0]
                    rv.append(to_filter[x])
        return rv, proj_dict

