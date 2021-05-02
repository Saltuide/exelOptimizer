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
                    value = cell.value.replace('_', '$%%$')
                    strings = re.sub('\W', ' ', value).split()
                    delimiters = re.sub('\w', '', value).replace('$%%$', '_')
                except Exception:
                    continue

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

                sheet.cell(row=i, column=j).value = new_value


    def a(self):
        sheet = self.file.active
        self.sheet_spelling_lopper(sheet)
        self.file.save('Files/new-main.xlsx')
