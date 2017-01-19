from __future__ import print_function
import os

import PcBusinessType as pc
import LifeBusinessType as life
import HealthBusinessType as health

import logging
from AMB_defines import *

logger = logging.getLogger('twolane')


class MasterProspectList:
    """Class containing all PC, Life and Health classes,,, pretty much all the data"""
    def __init__(self):
        self.file_info = None
        self.business_types = {}
        pass

    def get_filenames(self, data_dir):
        logger.info("Enter")
        logger.info("DATAPATH: %s", data_dir)

        files = os.listdir(data_dir)
        file_names = []
        for file in files:
            logger.info("file: %s", file)
            if file.find(sch_d) != -1:
                file_names.append(file)
            elif file.find(assets) != -1:
                file_names.append(file)
            elif file.find(cash_flow) != -1:
                file_names.append(file)
            elif file.find(sis) != -1:
                file_names.append(file)
            elif file.find(sch_ba) != -1:
                file_names.append(file)
        logger.info("Leave")
        return file_names

    def get_file_dict(self, data_dir):
        logger.info("Enter")
        file_names = self.get_filenames(data_dir)
        file_dict = {}

        for file in file_names:
            idx = file.find('_')
            entity_id = file[:idx]
            if entity_id in file_dict:
                file_dict[entity_id].append(file)
            else:
                file_dict[entity_id] = [file]
        logger.info("Leave")
        return file_dict

    def get_data_from_files(self, data_dir):
        csv_data_dir = data_dir + "\\data"
        file_info = self.get_file_dict(csv_data_dir)
        for entity_id in file_info:
            if entity_id == pc_tag:
                self.business_types[entity_id] = pc.PcBusinessType()
            elif entity_id == life_tag:
                self.business_types[entity_id] = life.LifeBusinessType()
            elif entity_id == health_tag:
                self.business_types[entity_id] = health.HealthBusinessType()
            self.business_types[entity_id].process_csvs( csv_data_dir, file_info[entity_id])
            #        build_4_sheets(xl_wb, entity_id)
        pass
    def get_pc_companies(self):
        return set(self.business_types[pc_tag].companies)

    def get_data_from_db(self):
        pass

