import os
import csv
from pathlib import Path

_FILENAME = 'stock.csv'


class StockLoader:
    def __init__(self):
        self.stock_dict = {}
        self.load_csv()

    def load_csv(self):
        project_path = os.getcwd()
        file_path = Path(project_path) / 'stock' / _FILENAME
        with open(file_path, newline='', encoding="utf-8") as csvfile:
            self.stock_dict = {row[0] : row[1] for row in csv.reader(csvfile)}

    def get_stock_name(self, code):
        return self.stock_dict.get(code)
