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
    def __init__(self, template_obj, data_dir):
        self.file_dict = None
        self.template_obj = template_obj
        self.business_types = {}
        self.build_business_types(data_dir)
        pass

    def get_filenames(self, data_dir):
        logger.info("Enter")
        logger.info("DATAPATH: %s", data_dir)

        files = os.listdir(data_dir)
        file_names = []
        for name in files:
            logger.info("file: %s", name)
            file_names.append(name)
        logger.info("Leave")
        return file_names

    def get_file_dict(self, data_dir):
        logger.info("Enter")
        file_names = self.get_filenames(data_dir)
        file_dict = {}

        for name in file_names:
            idx = name.find('_')
            if idx == -1:
                continue
            entity_id = name[:idx]
            if entity_id in file_dict:
                file_dict[entity_id].append(name)
            else:
                file_dict[entity_id] = [name]
        logger.info("Leave")
        return file_dict

    def build_business_types(self, data_dir):
        csv_data_dir = data_dir + "\\data"
        file_dict = self.get_file_dict(csv_data_dir)
        for entity_id in file_dict:
            if entity_id == PC_tag:
                self.business_types[entity_id] = pc.PcBusinessType()
            elif entity_id == LIFE_tag:
                self.business_types[entity_id] = life.LifeBusinessType()
            elif entity_id == HEALTH_tag:
                self.business_types[entity_id] = health.HealthBusinessType()
            else:
                continue
            self.business_types[entity_id].convert_csvs_to_raw_df(csv_data_dir, file_dict[entity_id])
            self.business_types[entity_id].construct_data_cube(self.template_obj)
        pass

    def get_companies(self, tag):
        if tag in self.business_types:
            return set(self.business_types[tag].companies)
        else:
            return None

    def build_trio_scorecard(self):
        logger.info("Enter")
        pass
        logger.info("Leave")
        return

    def get_candidates(self, tag):
        companies = self.get_companies(tag)
        if companies is None:
            return None
        else:
            return self.get_arbitrary_subset(companies)

    def get_arbitrary_subset(self, companies):
        """ Testing purposes only """
        tmp_set = set()
        for i in range(0, 1):
            tmp_set.add(companies.pop())
        return tmp_set

    def get_data_from_db(self):
        pass
