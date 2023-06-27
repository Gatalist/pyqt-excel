
class ReadDocument:
    def get_all_sheets(self) -> None:
        self.object_sheet.clear()
        for sheet_name in self.work_book.sheetnames:
            self.object_sheet.addItem(sheet_name)

    def get_all_columns(self) -> None:
        self.object_columns.clear()
        list_columns = self.data_frame.columns.values.tolist()
        for column in list_columns:
            self.object_columns.addItem(column)

    # возвращаем список нумерованых строк с данными в указаной ячейке
    def generate_list_data_row(self) -> list:
        data_rows = []
        start_number_string = 2
        for row in self.data_frame[self.column_name]:
            row_to_line = []
            if type(row) == str:
                row_to_line.append(start_number_string)
                row_to_line.append(row)
                data_rows.append(row_to_line)
            start_number_string += 1
        return data_rows

    # Выводим данные ячейки
    # Аргумент "read_line" разобъет текст на строки по символу ";"
    def read_list_data_row(self, list_data: str, read_line: bool) -> str:
        number = 0
        if len(list_data) > 0:
            for number_string, text in list_data:              
                string = f'-----[ Строка: {number_string} ]-----\n'
                if read_line:
                    call_data = text.split(';')
                    for line in call_data:
                        string += f'{line}\n'
                else:
                    string += f'{text}\n'
                yield string +'\n'
                number += 1
        else:
            alert = "\n⛔️ Колонка пустая\n"
            yield alert
        yield f"\n✅ Прочитано: {number} строк\n"

    # проверяем строку на первую заглавную букву
    def checking_is_title(self, list_data) -> str:
        errors = 0
        for number_string, text in list_data:            
            string = f'-----[ Строка {number_string} ]-----\n'
            check = False
            txt = ''
            for line in text.split(';'):
                try:
                    first_leter = line[0]
                    if first_leter.istitle() == False and first_leter[0].isdigit() == False:
                        errors += 1
                        check = True
                        txt += f'{line}\n'
                except IndexError:
                    errors += 1
                    check = True
                    txt += f'{line}\n'
            if check:
                yield string + f"{txt}\n"
        yield f"\n⛔️ Ошибок -> {errors}\n"

    # ищим фрагмент текста в ячейке
    def search_text(self, list_data, search: str) -> None:
        result = 0
        for number_string, text in list_data:
            string = f'-----[ Строка {number_string} ]-----\n'
            if text.lower().find(search.lower()) != -1:
                line = text.split(';')
                line_new = "\n".join(line)
                string += f'{line_new}\n'
                result += 1
                yield f"{string}\n"
        yield f"\n✅ Найдено совпадений -> {result}\n"

    def get_unique_strings(self, list_data):
        unique_elem = []
        for number_string, text in list_data:
            string = f'-----[ Строка {number_string} ]-----\n'
            for line in text.split(';'):
                if line not in unique_elem:
                    unique_elem.append(line)
                    yield f"{line}\n"
        yield f"\n✅ Найдено -> {len(unique_elem)}\n"


class ChangeDocument:
    # добавление фрагмента текста в начало ячейки если она не пустая
    def add_text_to_cell_start(self, cell_past: str, text: str):
        result = 0
        for number_string in self.list_len_string:
            cell_past_obj = self.get_cell_obj(cell_past, number_string)
            if cell_past_obj.value is not None:
                new_text = text + cell_past_obj.value
                self.save_result_in_cell(cell_past_obj, new_text)
                yield f'[{number_string}] : {new_text}\n'
                result += 1
        yield f"✅ Текст добавлен в начало строки в ячейку [ {cell_past} ] - {result}\n"

    # вырезаем весь текст с одной ячееки и добавляем в другую ячейку
    def move_text_to_another_cell(self, cell_move: str, cell_past: str) -> None:
        result = 0
        for number_string in self.list_len_string:
            cell_move_obj = self.get_cell_obj(cell_move, number_string)
            # вырезаем данные с ячейки если она не пустая
            if cell_move_obj.value is not None:
                cell_past_obj = self.get_cell_obj(cell_past, number_string)
                # вставляем данные в другую ячейку
                self.add_text_to_cell(cell_past_obj, text=cell_move_obj.value)
                yield f"[ {number_string} ] : {cell_move_obj.value}\n"
                # очищаем ячейку откуда копируем текст
                self.add_text_to_cell(cell_move_obj, text=None)
                result += 1
        yield (f"✅ Текст удален с колонки [{cell_move}] и добавлен в колонку [{cell_past}] - {result}\n")

    # удалить фрагмент текста со всех ячейк в столбце
    def remove_text_from_cell(self, cell_remove: str, text: str) -> None:
        result = 0
        for number_string in self.list_len_string:
            cell_remove_obj = self.get_cell_obj(cell_remove, number_string)
            if cell_remove_obj.value is not None:
                new_data = cell_remove_obj.value.replace(text, '')
                self.add_text_to_cell(cell_remove_obj, None)
                self.add_text_to_cell(cell_remove_obj, new_data)
                yield f"[ {number_string} ] : {new_data}\n"
                result += 1
        yield f"✅ Текст удален с [{cell_remove}] - {result}\n"

    # добавить текст перед текстом у весь столбц
    def add_text_after_text_cell(self, cell_past: str, after_text: str, text_past: str) -> None:
        result = 0
        for number_string in self.list_len_string:
            cell_past_obj = self.get_cell_obj(cell_past, number_string)
            if cell_past_obj.value is not None:
                new_data = cell_past_obj.value.replace(after_text, text_past)
                self.add_text_to_cell(cell_past_obj, None)
                self.add_text_to_cell(cell_past_obj, new_data)
                yield f"[ {number_string} ] : {new_data}\n"
                result += 1
        yield f"✅ Текст удален с [{cell_past}] - {result}\n"
    
    # добавление фрагмента текста в каждую ячейку
    def add_text_to_column(self, cell_past: str, text: str) -> None:
        result = 0
        for number_string in self.list_len_string:
            cell_move_obj = self.get_cell_obj(cell_past, number_string)
            self.add_text_to_cell(cell_move_obj, text)
            yield f'[{number_string}] : {text}\n'
            result += 1
        yield f"✅ Текст добавлен в колонку [ {cell_past} ] - {result}\n"

    # обьеденяем ячейки в одну
    def join_columns_text(self, save_column: str, join_columns: list, join_separator: str, end_text: str) -> None:
        for number_string in self.document.list_len_string:
            new_list_join = [col_row + str(number_string) for col_row in join_columns]
            new_data = []

            for column in new_list_join:
                old_data = self.get_cell_obj(column).value.replace(end_text, '').strip()
                new_data.append(old_data)
            
            new_text = join_separator.join(new_data)
            if new_text:
                new_text = new_text + ' ' + end_text
            
            cell_save = self.get_cell_obj(save_column, number_string)
            self.add_text_to_cell(cell_save, new_text)

    # удаяем текст поиска с ячейки и добавляем в другую ячейку
    def serch_move_text_to_another_cell(self, cell_move: str, cell_past: str, method_remove: str, search: list) -> None:
        result = 0
        for number_string in self.list_len_string:
            cell_move_obj = self.get_cell_obj(cell_move, number_string)
            search_text_lower = [word.lower() for word in search]
            
            if method_remove == 'str' and cell_move_obj.value is not None:
                current_text = cell_move_obj.value
                cell_past_list_new_text = []
                string = f'-----[ Line  {number_string} ]-----\n'
                chack = False
                # перебераем список совпадений
                for search_word in search_text_lower:
                    # перебераем список строк
                    for line in current_text.split(';'):
                        if line.lower().find(search_word) != -1:
                            if line not in cell_past_list_new_text:
                                cell_past_list_new_text.append(line)
                            # save text in current cell
                            curr_text = current_text.replace(line, '')
                            self.add_text_to_cell(cell_move_obj, None)
                            self.add_text_to_cell(cell_move_obj, curr_text)
                            chack = True

                    # добавление фрагмента текста в ячейку
                    cell_past_txt_add = ";".join(cell_past_list_new_text)
                    cell_past_obj = self.get_cell_obj(cell_past, number_string)
                    self.add_text_to_cell(cell_past_obj, cell_past_txt_add)

                if chack:
                    result += 1
                    yield f"{string + search_word}\n\n"

            if method_remove == 'txt':
                if cell_move_obj.value is not None and cell_move_obj.value.lower().find(search_text_lower[0]) != -1:
                    print(f'\n\n-----[ Line  {number_string} ]-----')
                    print(cell_move_obj.value)

                    # удалить фрагмент текста в ячейке
                    cell_move_txt = cell_move_obj.value.replace(search, '')
                    self.add_text_to_cell(cell_move_obj, cell_move_txt)

                    # добавление фрагмента текста в ячейку
                    cell_past_obj = self.get_cell_obj(cell_past, number_string)
                    self.add_text_to_cell(cell_past_obj, search_text_lower[0])
        yield f"✅ Изменено {result} строк\n\n"



class Document(ReadDocument, ChangeDocument):
    read_document: str = None # ссылка на документ
    data_frame: object = None  # получаем data_frame документа в pandas
    work_book: object = None  # открываем документ в openpyxl
    work_sheet: object = None  # получаем рабочий лист в документе
    column_name: str = None # имя колонки
    len_strings: int = None  # получаем список строк в документе
    start_number_string: int = 2 # с какой строки начинать читать документ
    list_len_string: list = None # список строк
    # current_method_work = None

    symbols = [';;', '; ;', '  ', '   ']

    # делаем первую букву каждой строки заглавной
    def upper_first_letter_in_text(self, text: str) -> str:
        split_text = text.split(';')
        new_list = []
        for line in split_text:
            capitalized = line[0:1].upper() + line[1:]
            new_list.append(capitalized)
        return ';'.join(new_list)
    
    # заменяем сымволы в строке
    def replace_symbol(self, text: str) -> str:
        for symbol in self.symbols:
            text.replace(symbol, ';')
        text.strip()

        if len(text) > 0:
            if text[0] == ';':
                text = text[1:]
        if len(text) > 0:
            if text[-1] == ';':
                text = text[:-1]
        return text.strip()

    # получаем обьект ячейки по координатам: например 'AC4'
    def get_cell_obj(self, cell_letter, cell_number) -> object:
        cell = f'{cell_letter}{cell_number}'
        return self.work_sheet[cell]    

    # сохраняем результат в ячейку
    def save_result_in_cell(self, cell_object: object, text: str) -> object:
        self.work_sheet[cell_object.coordinate] = text
        return self.read_document
    
    # добавление фрагмента текста в ячейку
    def add_text_to_cell(self, cell_object: object, text: str) -> object:
        if text:
            cell_object_value = cell_object.value
            if cell_object_value is not None:
                cell_object_value = f'{cell_object_value};{text}'
            else:
                cell_object_value = text

            clear_txt = self.replace_symbol(cell_object_value)
            upper_first_letter = self.upper_first_letter_in_text(clear_txt)
            self.save_result_in_cell(cell_object, upper_first_letter)
        else:
            self.save_result_in_cell(cell_object, None)
        return self.read_document
