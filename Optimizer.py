import re

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet, Cell
from spellchecker import SpellChecker

UNDEFINED_CELL = 'NA'


# noinspection PyMethodMayBeStatic
class Optimizer:

    def __init__(self, file: Workbook):
        self.file = file
        self.spell_checker = SpellChecker(language='ru')

    def strip(self, cell: Cell) -> None:
        """
        Убирает пробелы по краям

        :param cell: Ячейка для обработки
        :return: None
        """
        if isinstance(cell.value, str):
            value = cell.value
            cell.value = value.strip()

    def set_cell_not_null(self, cell: Cell) -> None:
        """
        Присваивает обозначение для пустой ячейки

        :param cell: Пустая ячейка
        :return: None
        """
        cell.value = UNDEFINED_CELL

    def is_string(self, value: str) -> bool:
        """
        Проверяет, является ли значение ячейки строкой

        :param value: Строка для проверки
        :return: True, если это строка
        """
        if value.isalpha():
            return True
        return False

    def check_spelling(self, value: str) -> str:
        """
        Исправляет синтаксические ошибки в строке

        :param value: Строка для исправления ошибок
        :return: Исправленная строка
        """
        return self.spell_checker.correction(value)

    def create_new_cell_value(self, strings: list, delimiters: str) -> str:
        """
        Собирает новое исправленное значение ячейки

        :param strings: Массив строковых значений
        :param delimiters: Строка с исходными разделителями
        :return: Новое значение ячейки
        """
        new_value = ''
        for counter, string in enumerate(strings):
            if self.is_string(string):
                new_value += self.check_spelling(string)
                if counter < len(delimiters):
                    new_value += delimiters[counter]
            else:
                new_value += string
                if counter < len(delimiters):
                    new_value += delimiters[counter]

        return new_value

    def sheet_spelling_lopper(self, sheet: Worksheet) -> None:
        """
        Проходит по листу и исправляет базовые ошибки

        :param sheet: Текущий лист
        :return: None
        """
        for i in range(2, sheet.max_row + 1):
            for j in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=i, column=j)

                self.strip(cell)

                if cell.value is None or cell.value == '':
                    self.set_cell_not_null(cell)
                    continue

                try:
                    value = cell.value.replace('_', '$%%$')
                    strings = re.sub('\W', ' ', value).split()
                    delimiters = re.sub('\w', '', value).replace('$%%$', '_')
                except Exception:
                    continue

                cell.value = self.create_new_cell_value(strings, delimiters)

    def a(self):
        sheet = self.file.active
        self.sheet_spelling_lopper(sheet)
        self.file.save('Files/new-main.xlsx')
