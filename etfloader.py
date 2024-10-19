import os
import openpyxl
from pathlib import Path

_FILENAME = 'etfs.xlsx'


class ETFLoader:
    def __init__(self):
        self.etf_issuers = []
        self.etf_codes = []
        self.etf_names = []
        self.load_xlsx()

    def load_xlsx(self):
        project_path = os.getcwd()
        path = Path(project_path) / 'etf_issuer'
        os.chdir(path)
        workbook = openpyxl.load_workbook(_FILENAME)
        for sheet_name in workbook.sheetnames:
            self.etf_issuers.append(sheet_name)
            sheet = workbook[sheet_name]
            codes = []
            names = []
            for row in sheet.iter_rows(values_only=True):
                codes.append(row[0])
                names.append(row[1])
            self.etf_codes.append(codes)
            self.etf_names.append(names)
        os.chdir(project_path)

    def get_title(self, index):
        titles = []
        size = len(self.etf_codes[index])
        for i in range(size):
            title = self.etf_codes[index][i] + ' ' + self.etf_names[index][i]
            titles.append(title)
        return titles

    def has_etf(self, etf):
        for etfs in self.etf_codes:
            if etf in etfs:
                return True
        return False
