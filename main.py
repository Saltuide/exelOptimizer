from spellchecker import SpellChecker

import openpyxl

from Optimizer import Optimizer

wb = openpyxl.load_workbook('Files/main.xlsx', data_only=True)


optimizer = Optimizer(wb)
optimizer.a()

