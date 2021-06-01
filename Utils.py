from spellchecker import SpellChecker
from openpyxl.styles.borders import Border, Side


class Utils:

    @staticmethod
    def is_string(value: str) -> bool:
        """
        Проверяет, является ли значение ячейки строкой

        :param value: Строка для проверки
        :return: True, если это строка
        """
        if value.isalpha():
            return True
        return False

    @staticmethod
    def check_spelling(spell_checker: SpellChecker, value: str) -> str:
        """
        Исправляет синтаксические ошибки в строке

        :param spell_checker:
        :param value: Строка для исправления ошибок
        :return: Исправленная строка
        """
        if len(value) < 3:
            return value

        return spell_checker.correction(value)

    @staticmethod
    def create_new_cell_value(check_spelling: SpellChecker, strings: list, delimiters: str) -> str:
        """
        Собирает новое исправленное значение ячейки

        :param check_spelling:
        :param strings: Массив строковых значений
        :param delimiters: Строка с исходными разделителями
        :return: Новое значение ячейки
        """
        new_value = ''
        for counter, string in enumerate(strings):
            if Utils.is_string(string):
                new_value += Utils.check_spelling(check_spelling, string)
                if counter < len(delimiters):
                    new_value += delimiters[counter]
            else:
                new_value += string
                if counter < len(delimiters):
                    new_value += delimiters[counter]

        return new_value

    @staticmethod
    def get_thin_border():
        return Border(left=Side(style='thin'),
                      right=Side(style='thin'),
                      top=Side(style='thin'),
                      bottom=Side(style='thin'))
