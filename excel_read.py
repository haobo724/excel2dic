import os
from pathlib import Path

import numpy as np
import pandas as pd

# Define input out
path = Path(os.path.dirname(__file__))
path_in = os.path.join(path, "FP")
path_out = os.path.join(path, "Results")
if not os.path.exists(path_out):
    os.makedirs(path_out)

class Excel2Dic:
    def __init__(self, csv_name='People_Asia'):
        self.xls_table_key = \
            pd.read_excel(os.path.join(path_in, f"{csv_name}.xlsx"), sheet_name=None, skiprows=0).keys()
        t = list(self.xls_table_key)[0]
        self.xls_table = pd.read_excel(os.path.join(path_in, f"{csv_name}.xlsx"), sheet_name=None, skiprows=0)[t]

        self.contoury_name = list(self.xls_table.iloc[:, 0])
        self.dict = {}
        self.years = [0]
        self._get_years()
        self._generate_dic()
        self.name = csv_name
    def _check_name_validation(self, name: str):
        if name in self.contoury_name:
            return True
        else:
            print(f'{name} is not in {self.name}table')
            return False

    def _check_year_validation(self, year: str):
        if int(year) in self.years:
            return True
        else:
            return 0

    def _generate_dic(self):
        contoury_dict = {}
        conter = 1
        for y in self.years:
            for i in range(self.xls_table.__len__()):
                value = list(self.xls_table.iloc[i])[conter]
                try:
                    value = float(value)
                except ValueError:
                    value=0

                contoury_dict.setdefault(self.contoury_name[i], value)

            conter += 1
            self.dict.setdefault(str(y), contoury_dict)
            contoury_dict = {}

    def get_value(self, year: str, name: str):
        if not self._check_year_validation(year):
            return 0
        if self._check_name_validation(name):
            value = self.dict[year][name]
            if np.isnan(value):
                return 0
            return value
        else:
            return 0

    def get_year_value(self, year: str):
        if not self._check_year_validation(year):
            return 0

        return self.dict[year]

    def _get_years(self):
        self.years = list(self.xls_table.iloc[:0])[1:]


