import re

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet, Cell
from spellchecker import SpellChecker


class Optimizer:

    def __init__(self, file: Workbook):
        self.file = file
        self.spell_checker = SpellChecker(language='ru')

    def is_string(self, value: str) -> bool:
        if value.isalpha():
            return True
        return False

    def check_spelling(self, value: str) -> str:
        return self.spell_checker.correction(value)

    def sheet_spelling_lopper(self, sheet: Worksheet) -> None:
        for i in range(1, sheet.max_row + 1):
            for j in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=i, column=j)
                if cell.value is None:
                    continue

                try:
                    strings = re.sub('\W', ' ', cell.value).split()
                except TypeError:
                    continue

                new_values = []
                for string in strings:
                    if self.is_string(string):
                        new_values.append(self.check_spelling(string))
                    else:
                        new_values.append(string)

                sheet.cell(row=i, column=j).value = ' '.join(new_values)


    def a(self):
        sheet = self.file.active
        self.sheet_spelling_lopper(sheet)
        self.file.save('Files/new-main.xlsx')
