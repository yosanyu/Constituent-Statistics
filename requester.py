import xlsxwriter
import requests
from bs4 import BeautifulSoup
import re
import time
import os

_FILENAME = 'ETF_Statistics.xlsx'
_ETF_HEADERS = ['股票代碼', '股票名稱', '權重']
_STATISTICS_HEADERS = ['代號', '名稱', '次數', '百分比', '最大權重ETF',
                       '最大權重', '最小權重ETF', '最小權重', '平均權重']

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
}


def sleep():
    time.sleep(3)


def avoid_weight_except(weight):
    try:
        float_weight = float(weight)
    except ValueError:
        return None
    return float_weight


class ETFRequester:
    def __init__(self, etfs):
        self.etfs = etfs
        self.callback_print_message = None
        self.callback_get_stock = None
        self.callback_end = None
        self.stock_dict = {}
        self.stock_max_weight_dict = {}
        self.stock_min_weight_dict = {}
        self.stock_average_weight_dict = {}
        self.base_url = 'https://www.moneydj.com/ETF/X/Basic/Basic0007B.xdjhtm?etfid='
        self.len = len(self.etfs)
        self.string_len = str(self.len)
        self.workbook = xlsxwriter.Workbook(_FILENAME)
        self.bold_format = self.workbook.add_format({'bold': False})

    def start_request(self):
        self.write_etf_data()
        self.write_statistics_data()
        self.workbook.close()
        self.clear()
        self.callback_print_message(_FILENAME + '文件已保存\n')
        self.callback_print_message('自動開啟文件' + _FILENAME + '中...\n')
        self.callback_end(True)
        try:
            os.startfile(_FILENAME)
        except Exception as e:
            self.callback_print_message('自動開啟文件' + _FILENAME + '失敗\n')

    def set_weight(self, key, etf, weight):
        self.update_max_weight(key, etf, weight)
        self.update_min_weight(key, etf, weight)
        self.update_average_weight(key, weight)

    def update_max_weight(self, key, etf, weight):
        float_weight = avoid_weight_except(weight)
        if float_weight:
            _, max_weight = self.stock_max_weight_dict.get(key, (None, float('-1.0')))
            if float_weight > max_weight:
                self.stock_max_weight_dict[key] = (etf, float_weight)

    def update_min_weight(self, key, etf, weight):
        float_weight = avoid_weight_except(weight)
        if float_weight:
            _, min_weight = self.stock_min_weight_dict.get(key, (None, float('inf')))
            if float_weight < min_weight:
                self.stock_min_weight_dict[key] = (etf, float_weight)

    def update_average_weight(self, key, weight):
        float_weight = avoid_weight_except(weight)
        if float_weight:
            total_weight, count = self.stock_average_weight_dict.get(key, (0.0, 0))
            self.stock_average_weight_dict[key] = (total_weight + float_weight, count + 1)

    def clear(self):
        self.stock_dict = {}
        self.stock_max_weight_dict = {}
        self.stock_min_weight_dict = {}
        self.stock_average_weight_dict = {}

    def write_etf_data(self):
        for index, etf in enumerate(self.etfs, start=1):
            self.print_etf_message(etf, index)
            soup = self.request_etf_data(etf)
            if not soup:
                sleep()
                continue
            weights = soup.findAll('td', 'col06')
            worksheet = self.workbook.add_worksheet(etf)
            self.write_worksheet(worksheet, 0, _ETF_HEADERS)
            for stock_index, stock in enumerate(soup.findAll('td', 'col05')):
                content = stock.find('a')
                text = content.get('href')
                match = re.findall('[0-9]+', text)
                # 避免抓取有誤而超出邊界
                if len(match) > 2:
                    key = match[1]
                    self.set_stock_times(key)
                    weight = weights[stock_index].get_text()
                    self.set_weight(key, etf, weight)
                    name = self.callback_get_stock(key)
                    cols = [key, name, weight]
                    for col, value in enumerate(cols):
                        worksheet.write(stock_index+1, col, value, self.bold_format)
            sleep()

    def print_etf_message(self, etf, index):
        message = f'正在請求{etf}的成分股中... 進度 {index}/{self.string_len}\n'
        self.callback_print_message(message)

    def request_etf_data(self, etf):
        url = self.base_url + etf + '.TW'
        try:
            request = requests.get(url, headers=headers)
        except (requests.exceptions.RequestException, Exception) as e:
            self.callback_print_message('發生異常, 請求' + etf + '資料失敗\n')
            return None
        request.encoding = 'utf-8'
        soup = BeautifulSoup(request.text, "html.parser")
        return soup

    def write_worksheet(self, worksheet, row, data):
        for col, value in enumerate(data):
            worksheet.write(row, col, value, self.bold_format)

    def set_stock_times(self, key):
        # 個股代號為4碼
        if len(key) == 4:
            self.stock_dict[key] = self.stock_dict.get(key, 0) + 1

    def write_statistics_data(self):
        worksheet = self.workbook.add_worksheet('統計')
        # 根據value做降冪排序
        sort_desc = dict(sorted(self.stock_dict.items(), key=lambda x: x[1], reverse=True))
        self.write_worksheet(worksheet, 0, _STATISTICS_HEADERS)
        row = 1
        for key, value in sort_desc.items():
            # 未重疊不記錄
            if value == 1:
                row += 1
                continue
            count = f'{value}/{self.string_len}'
            percentage = f'{round(value / self.len * 100, 2)}%'
            name = self.callback_get_stock(key)
            max_weight, min_weight = self.stock_max_weight_dict.get(key), self.stock_min_weight_dict.get(key)
            average_weight = self.stock_average_weight_dict.get(key)[0] / self.stock_average_weight_dict.get(key)[1]
            data = [key, name, count, percentage, max_weight[0], max_weight[1],
                    min_weight[0], min_weight[1], average_weight]
            self.write_worksheet(worksheet, row, data)
            row += 1
