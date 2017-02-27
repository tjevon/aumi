from __future__ import print_function

import xlwings as xw

from AMB_defines import *

logger = logging.getLogger('twolane')


class TemplateWorkBook(object):

    template_row_labels = {
        SI01_tag: ('A', 3, 154),
        E07_tag: ('A',  3, 51),
        Assets_tag: ('A', 3, 60),
        CashFlow_tag: ('A', 3, 48),
        SI05_07_tag: ('A', 3, 83),
        SoI_tag: ('A', 3, 79),
        SoO_tag: ('A', 3, 80),
        SoR_tag: ('A', 3, 72),
        IRIS1_tag: ('A', 3, 30),
        IRIS2_tag: ('A', 3, 28),
        Liab1_tag: ('A', 3, 52),
        Liab2_tag: ('A', 3, 67),
        Liab3_tag: ('A', 3, 41),
        CR_tag: ('A', 3, 4)
    }

    # fid located in first column, display filter in second
    template_BT_cols = {
        PC_tag: ('B', 'C', 'L'),
        LIFE_tag: ('D', 'E', 'M'),
        HEALTH_tag: ('F', 'G', 'N')
    }

    def __init__(self, data_dir):
        self.xl_wb = self.get_template_workbook(data_dir)

    @staticmethod
    def get_template_workbook(data_dir):
        template_data_dir = data_dir + "\\templates"
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

    def get_full_fid_list(self, tag, bt, y_or_q=YEARLY_IDX):
        template_sheet = self.get_template_sheet(tag)
        range_info = self.template_row_labels[tag]

        a_cols = []
        for i in range(range_info[1], range_info[2]+1):
            a_cols.append(str(range_info[0]) + str(i))

        fid_col = self.template_BT_cols[bt.get_bt_tag()][y_or_q]
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

    def get_display_fid_list(self, tag, bt, y_or_q):
        fid_list = self.get_full_fid_list(tag, bt, y_or_q)[0]
        rv = self.filter_list(tag, bt.get_bt_tag(), fid_list, 1)
        return rv

    def get_row_labels(self, tag, bt_tag):
        template_sheet = self.get_template_sheet(tag)
        range_info = self.template_row_labels[tag]
        cells = str(range_info[0]) + str(range_info[1]) + ':' + str(range_info[0]) + str(range_info[2])
        row_labels = template_sheet.range(cells).options(ndim=2).value

        row_labels = self.filter_list(tag, bt_tag, row_labels, 0)
        return row_labels

    def filter_list(self, tag, bt_tag, to_filter, offset):
        template_sheet = self.get_template_sheet(tag)
        range_info = self.template_row_labels[tag]
        display_filter_col = self.template_BT_cols[bt_tag][DISPLAY_IDX]
        cells = display_filter_col + str(range_info[1]+offset) + ':' + display_filter_col + str(range_info[2])
        filters = template_sheet.range(cells).options(ndim=2).value
        rv = []
        for filter_cell, x in zip(filters, range(len(filters))):
            if filter_cell[0] == 1:
                continue
            rv.append(to_filter[x])
        return rv
