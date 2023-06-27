from dataclasses import dataclass


@dataclass
class Variable:
    object_edition_read_text: str = "Читать"
    object_edition_change_text: str = "Редактировать"
    
    object_option_read_col: str = 'Прочитать колонку'
    object_option_check_start_end_str: str = 'Проверить начало и конец строки'  
    object_search_text_str: str = 'Поиск текста'
    object_unique_strings_str: str = 'Уникальные строки'
    
    add_text_to_cell_start_txt: str = 'Добавление текста в начало (не пустая строка)'
    add_text_to_cell_end_txt: str = 'Добавление текста в конец (везде)'
    remove_text_from_cell_txt: str = 'Удалить указаный текст с столбца'
    move_text_to_another_cell_txt: str = 'Вызерать текст в другую ячейку (весь)'
    serch_move_text_to_another_cell_txt: str = 'Вызерать совпадение текста в другую ячейку (поиск)'
    add_text_after_text_cell_txt: str = 'Добавить текст перед текстом'

    explorer_open_doc_txt = 'Open File'
    explorer_save_doc_txt = 'Save File'

    save_doc_text = '\n\n💾 Save [ Document ]\n'
    faled_doc_text = '\n\n⛔️ Failed to save [ Document ]\n'

    format_document = 'Excel (*.xlsx);;Excel (*.xls)'

class LogicsReadDocument(Variable):
    # методы над документом
    def hiden_frame_col_1(self):
        self.frame_object_search_text.setVisible(False)

    
    def check_select_option_read(self):
        print("check select item")
        self.hiden_frame_col_1()
        if self.object_option_read.currentText() == self.object_search_text_str:
            self.frame_object_search_text.setVisible(True)


    def select_option_read(self):
        print("Читать")
        self.object_option_read.addItem(self.object_option_read_col)
        self.object_option_read.addItem(self.object_option_check_start_end_str)
        self.object_option_read.addItem(self.object_search_text_str)
        self.object_option_read.addItem(self.object_unique_strings_str)

    # запуск программы
    def btn_start_read(self):
        self.work_sheet = self.work_book[self.object_sheet.currentText()] # получаем рабочий лист в документе
        self.column_name = self.object_columns.currentText() # получаем текущую колонку для работы
        self.output_widget.clear() # очищаем текстовое поле
        self.object_option_read.currentText()  # получаем выбраный метод работы с документом
        list_data = self.generate_list_data_row()
        method = self.object_option_read.currentText()
        
        if method == self.object_option_read_col:
            # Выводим данные ячейки
            read = self.read_list_data_row(list_data, read_line=True)
            for string in read:
                self.output_widget.appendPlainText(string)
        
        if method == self.object_option_check_start_end_str:
            # проверяем строку на первую заглавную букву
            read = self.checking_is_title(list_data)
            for string in read:
                self.output_widget.appendPlainText(string)

        if method == self.object_search_text_str:
            # ищим фрагмент текста в ячейке
            obj_text = self.object_search_text.toPlainText()  # получаем текст для поиска с поле для ввода
            read = self.search_text(list_data=list_data, search=obj_text)
            for string in read:
                self.output_widget.appendPlainText(string)

        if method == self.object_unique_strings_str:
            read = self.get_unique_strings(list_data)
            for string in read:
                self.output_widget.appendPlainText(string)


class LogicsChangeDocument(Variable):
    def hiden_frame_col(self):
        self.frame_copy.setVisible(False)
        self.frame_past.setVisible(False)
        self.frame_object_search_text_2.setVisible(True)
        self.label_6.setVisible(False)
        self.label_10.setVisible(False)
        self.label_11.setVisible(False)
        self.object_after_text.setVisible(False)

    def check_select_option_change(self):
        print("check select item")
        self.hiden_frame_col()
        if self.object_option_change.currentText() == self.add_text_to_cell_start_txt:
            self.frame_past.setVisible(True)
        if self.object_option_change.currentText() == self.add_text_to_cell_end_txt:
            self.frame_past.setVisible(True)
        if self.object_option_change.currentText() == self.remove_text_from_cell_txt:
            self.frame_copy.setVisible(True)
        if self.object_option_change.currentText() == self.move_text_to_another_cell_txt:
            self.frame_copy.setVisible(True)
            self.frame_past.setVisible(True)
            self.frame_object_search_text_2.setVisible(False)
        if self.object_option_change.currentText() == self.serch_move_text_to_another_cell_txt:
            self.frame_copy.setVisible(True)
            self.frame_past.setVisible(True)
            self.label_6.setVisible(True)
        if self.object_option_change.currentText() == self.add_text_after_text_cell_txt:
            self.object_after_text.setVisible(True)
            self.frame_past.setVisible(True)
            self.label_6.setVisible(True)
            self.label_10.setVisible(True)
            self.label_11.setVisible(True)
        

    # методы над документом
    def select_option_change(self):
        print("Редактировать")
        self.hiden_frame_col()
        self.object_option_change.addItem(self.add_text_to_cell_start_txt)
        self.object_option_change.addItem(self.add_text_to_cell_end_txt)
        self.object_option_change.addItem(self.remove_text_from_cell_txt)
        self.object_option_change.addItem(self.move_text_to_another_cell_txt)
        self.object_option_change.addItem(self.serch_move_text_to_another_cell_txt)
        self.object_option_change.addItem(self.add_text_after_text_cell_txt)

    # запуск программы
    def btn_start_change(self):
        self.work_sheet = self.work_book[self.object_sheet.currentText()] # получаем рабочий лист в документе
        self.column_name = self.object_columns.currentText() # получаем текущую колонку для работы
        self.output_widget.clear() # очищаем текстовое поле
        method = self.object_option_change.currentText()  # получаем выбраный метод работы с документом
        
        #+ добавление фрагмента текста в начало ячейки если она не пустая
        if method == self.add_text_to_cell_start_txt:
            past_text = self.object_past.toPlainText()
            obj_search_text = self.object_search_text_2.toPlainText()
            write = self.add_text_to_cell_start(cell_past=past_text, text=obj_search_text)
            for string in write:
                self.output_widget.appendPlainText(string)
                print(string)

        #+ добавление фрагмента текста в каждую ячейку
        if method == self.add_text_to_cell_end_txt:
            past_text = self.object_past.toPlainText()
            obj_text = self.object_search_text_2.toPlainText()
            write = self.add_text_to_column(cell_past=past_text, text=obj_text)
            for string in write:
                self.output_widget.appendPlainText(string)
                print(string)

        #+ удалить фрагмент текста со всех ячейк в столбце
        if method == self.remove_text_from_cell_txt:
            remove_text = self.object_copy.toPlainText()
            obj_text = self.object_search_text_2.toPlainText()
            remove = self.remove_text_from_cell(cell_remove=remove_text, text=obj_text)
            for string in remove:
                self.output_widget.appendPlainText(string)
                print(string)

        # --- удалить фрагмент текста со всех ячейк в столбце
        if method == self.add_text_after_text_cell_txt:
            cell_past = self.object_past.toPlainText()
            after_text = self.object_after_text.toPlainText()
            text_past = self.object_search_text_2.toPlainText()
            remove = self.add_text_after_text_cell(cell_past=cell_past, after_text=after_text, text_past=text_past)
            for string in remove:
                self.output_widget.appendPlainText(string)
                print(string)

        #+ вырезаем весь текст с одной ячееки и добавляем в другую ячейку
        if method == self.move_text_to_another_cell_txt:
            copy = self.object_copy.toPlainText()
            past = self.object_past.toPlainText()
            remove = self.move_text_to_another_cell(cell_move=copy, cell_past=past)
            for string in remove:
                self.output_widget.appendPlainText(string)
                print(string)

        # удаяем текст поиска с ячейки и добавляем в другую ячейку
        if method == self.serch_move_text_to_another_cell_txt:
            copy = self.object_copy.toPlainText()
            past = self.object_past.toPlainText()
            obj_text = self.object_search_text_2.toPlainText()
            search_text = [text for text in obj_text.split('::')]

            remove = self.serch_move_text_to_another_cell(cell_move=copy, cell_past=past, method_remove='str', search=search_text)
            for string in remove:
                self.output_widget.appendPlainText(string)
                print(string)
                
        # обьяденяем данные столбцов в один столбец
            # self.join_columns_text(save_column='AH', join_columns=['AE', 'AF', 'AG'], join_separator=' x ', end_text='см')