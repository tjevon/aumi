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
    def __init__(self, template_wb):
        self.file_dict = None
        self.template_wb = template_wb
        self.business_types = {}
        pass

    def get_filenames(self, data_dir):
        logger.info("Enter")
        logger.info("DATAPATH: %s", data_dir)

        files = os.listdir(data_dir)
        file_names = []
        for file in files:
            logger.info("file: %s", file)
            if file.find(Assets_files) != -1:
                file_names.append(file)
            elif file.find(CashFlow_files) != -1:
                file_names.append(file)
            elif file.find(E07_files) != -1:
                file_names.append(file)
            elif file.find(E10_files) != -1:
                file_names.append(file)
            elif file.find(LiabSurp_files) != -1:
                file_names.append(file)
            elif file.find(SI01_files) != -1:
                file_names.append(file)
            elif file.find(SI05_07_files) != -1:
                file_names.append(file)
            elif file.find(SI08_09_files) != -1:
                file_names.append(file)
            elif file.find(SoI_files) != -1:
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
        file_dict = self.get_file_dict(csv_data_dir)
        for entity_id in file_dict:
            if entity_id == pc_tag:
                self.business_types[entity_id] = pc.PcBusinessType()
                self.business_types[entity_id].process_csvs( csv_data_dir, file_dict[entity_id])
                self.business_types[entity_id].construct_data_cubes(self.template_wb)
            # elif entity_id == life_tag:
            #     self.business_types[entity_id] = life.LifeBusinessType()
            #     self.business_types[entity_id].process_csvs( csv_data_dir, file_dict[entity_id])
            # elif entity_id == health_tag:
            #     self.business_types[entity_id] = health.HealthBusinessType()
            #     self.business_types[entity_id].process_csvs( csv_data_dir, file_dict[entity_id])
            # #        build_4_sheets(xl_wb, entity_id)
        pass

    def get_pc_companies(self):
        return set(self.business_types[pc_tag].companies)

    def get_data_from_db(self):
        pass

