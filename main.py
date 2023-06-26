import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMenu
from openpyxl import load_workbook
import pandas

from ui_form import Ui_MainWindow
from document import Document
from logica import LogicsReadDocument, LogicsChangeDocument


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow, Document, LogicsReadDocument, LogicsChangeDocument):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)
        self.setFixedSize(1000, 660)

        # привязываем события
        self.object_btn_start_read.clicked.connect(self.btn_start_read)
        self.object_btn_start_change.clicked.connect(self.btn_start_change)

        self.object_option_read.activated.connect(self.check_select_option_read)
        self.object_option_change.activated.connect(self.check_select_option_change)

        self.create_menu_bar()
        self.select_option_read()
        self.select_option_change()
        self.check_select_option_change()
        self.check_select_option_read()
        
    
    # создаем меню
    def create_menu_bar(self):
        menuBar = self.menuBar()
        # File menu
        fileMenu = QMenu("Файл", self)
        menuBar.addMenu(fileMenu)
        fileMenu.addAction("Открыть").triggered.connect(self.open_excel_file)
        fileMenu.addAction("Сохранить").triggered.connect(self.save_excel_file)


    # выбираем файл и добавляем настройки
    def open_excel_file(self):
        file = QtWidgets.QFileDialog.getOpenFileName(self, self.explorer_open_doc_txt, './', self.format_document)
        if file:
            self.read_document = file[0]
            if self.read_document:
                self.data_frame: object = pandas.read_excel(self.read_document) # получаем data_frame документа в pandas
                self.work_book: object = load_workbook(self.read_document)  # открываем документ в openpyxl
                self.get_all_sheets()  # получаем все листы в документе
                self.get_all_columns() # получаем все колонки в документе
                self.len_strings = len(self.data_frame.index) + 1 # получаем список строк в документе
                self.list_len_string = [string for string in range(2, self.len_strings + 1)] # список строк
                self.output_widget.appendPlainText(f"\n✅ {self.explorer_open_doc_txt} [ {file[0]} ]")

    # save new file
    def save_excel_file(self):
        name = QtWidgets.QFileDialog.getSaveFileName(self, self.explorer_save_doc_txt, './', self.format_document)
        string = ""
        try:
            self.work_book.save(name[0])
            string = self.save_doc_text
        except:
            string = self.faled_doc_text

        self.output_widget.appendPlainText(string)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()
