import etfloader
import requester
import stockloader
from PySide6.QtCore import QThread
from PySide6.QtCore import Signal
from PySide6.QtWidgets import QMainWindow
from PySide6.QtWidgets import QPushButton
from PySide6.QtWidgets import QComboBox
from PySide6.QtWidgets import QPlainTextEdit


_DIVIDEND_ETFS = ['0056', '00701', '00713', '00730', '00731', '00878', '00900',
                  '00907', '00915', '00918', '00919', '00927', '00929', '00930',
                  '00932', '00934', '00936', '00939', '00940', '00943', '00944',
                  '00946', '00961']


class WorkerThread(QThread):
    print_message = Signal(str)
    end = Signal(bool)

    def __init__(self, etf_requester, parent=None):
        self.etf_requester = etf_requester
        super().__init__(parent)

    def run(self):
        # 綁定回調函數
        self.etf_requester.callback_print_message = self.callback_print_message
        self.etf_requester.callback_end = self.callback_end
        self.etf_requester.start_request()

    def callback_print_message(self, text):
        self.print_message.emit(text)

    def callback_end(self, enable):
        self.end.emit(enable)


class MainWindow(QMainWindow):
    def __init__(self):
        self.etf_loader = etfloader.ETFLoader()
        self.stock_loader = stockloader.StockLoader()
        self.etfs = []
        self.plain_text = ''
        self.table = None
        self.button_add = None
        self.button_confirm = None
        self.button_clear = None
        self.button_dividends = None
        self.combobox_etf_issuer = None
        self.combobox_etf_title = None
        self.plain_text_edit = None
        self.worker_thread = None
        self.window_setting()
        self.init_widget()

    def window_setting(self):
        super(MainWindow, self).__init__()
        self.resize(1600, 900)
        self.setFixedSize(1600, 900)
        self.setWindowTitle("Constituent Statistics")

    def init_widget(self):
        self.init_button()
        self.init_combobox()
        self.init_plain_text_edit()

    def init_button(self):
        self.button_add = QPushButton(self)
        self.button_confirm = QPushButton(self)
        self.button_clear = QPushButton(self)
        self.button_dividends = QPushButton(self)
        self.button_add.setGeometry(1200, 50, 100, 50)
        self.button_confirm.setGeometry(1320, 50, 100, 50)
        self.button_clear.setGeometry(1440, 50, 100, 50)
        self.button_dividends.setGeometry(1200, 120, 200, 50)
        self.button_add.setText("新增")
        self.button_confirm.setText("確認")
        self.button_clear.setText("重置")
        self.button_dividends.setText('股息系列')
        font = self.button_add.font()
        font.setPointSize(20)
        self.button_add.setFont(font)
        self.button_confirm.setFont(font)
        self.button_clear.setFont(font)
        self.button_dividends.setFont(font)
        self.button_add.clicked.connect(self.on_button_add_clicked)
        self.button_confirm.clicked.connect(self.on_button_confirm_clicked)
        self.button_clear.clicked.connect(self.on_button_clear_clicked)
        self.button_dividends.clicked.connect(self.on_button_dividends_clicked)

    def init_combobox(self):
        self.combobox_etf_issuer = QComboBox(self)
        self.combobox_etf_title = QComboBox(self)
        self.combobox_etf_issuer.setGeometry(50, 50, 250, 50)
        self.combobox_etf_title.setGeometry(350, 50, 800, 50)
        self.combobox_etf_issuer.addItems(self.etf_loader.etf_issuers)
        self.combobox_etf_title.addItems(self.etf_loader.get_title(0))
        font = self.combobox_etf_issuer.font()
        font.setPointSize(20)
        self.combobox_etf_issuer.setFont(font)
        self.combobox_etf_title.setFont(font)
        self.combobox_etf_issuer.currentIndexChanged.connect(self.etf_issuer_changed)

    def init_plain_text_edit(self):
        self.plain_text_edit = QPlainTextEdit(self)
        self.plain_text_edit.setGeometry(100, 200, 1400, 600)
        font = self.plain_text_edit.font()
        font.setPointSize(30)
        self.plain_text_edit.setFont(font)
        self.update_plain_text()
        self.plain_text_edit.show()

    def set_button_is_enable(self, enable):
        buttons = [self.button_add, self.button_confirm,
                   self.button_clear, self.button_dividends]
        for button in buttons:
            button.setEnabled(enable)

    def on_button_add_clicked(self):
        etf_issuer_index = self.combobox_etf_issuer.currentIndex()
        etf_code_index = self.combobox_etf_title.currentIndex()
        etf_code = self.etf_loader.etf_codes[etf_issuer_index][etf_code_index]
        if etf_code not in self.etfs:
            self.etfs.append(etf_code)
            message = '已新增{}進入統計\n'.format(etf_code)
            self.add_plain_text(message)

    def on_button_confirm_clicked(self):
        if len(self.etfs) > 0:
            self.set_button_is_enable(False)
            etf_requester = requester.ETFRequester(self.etfs)
            # 綁定回調函數
            etf_requester.callback_get_stock = self.get_stock_name
            # 建立QT Thread
            self.worker_thread = WorkerThread(etf_requester)
            self.worker_thread.print_message.connect(self.add_plain_text)
            self.worker_thread.end.connect(self.set_button_is_enable)
            self.worker_thread.start()

    def on_button_clear_clicked(self):
        self.etfs = []
        self.clear_plain_text()

    def on_button_dividends_clicked(self):
        self.on_button_clear_clicked()
        self.set_button_is_enable(False)
        for etf in _DIVIDEND_ETFS:
            if etf not in self.etfs:
                self.etfs.append(etf)
                message = '已新增{}進入統計\n'.format(etf)
                self.add_plain_text(message)
        self.on_button_confirm_clicked()

    def etf_issuer_changed(self, index):
        self.combobox_etf_title.clear()
        self.combobox_etf_title.addItems(self.etf_loader.get_title(index))

    def add_plain_text(self, text):
        self.plain_text += text
        self.update_plain_text()

    def clear_plain_text(self):
        self.plain_text = ''
        self.update_plain_text()

    def update_plain_text(self):
        self.plain_text_edit.setPlainText(self.plain_text)
        scrollbar = self.plain_text_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def get_stock_name(self, key):
        return self.stock_loader.get_stock_name(key)