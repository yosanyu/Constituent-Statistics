import os
import csv
from pathlib import Path


class StockLoader:
    def __init__(self):
        self.stock_dict = {}
        self.load_csv()

    def load_csv(self):
        project_path = os.getcwd()
        path = Path(project_path) / 'stock'
        os.chdir(path)
        with open('stock.csv', newline='', encoding="utf-8") as csvfile:
            rows = csv.reader(csvfile)
            for row in rows:
                key, value = row[0], row[1]
                self.stock_dict[key] = value
        os.chdir(project_path)

    def get_stock_name(self, code):
        return self.stock_dict.get(code)
