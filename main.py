import re
import pandas as pd
import numpy as np
from skimage.measure import label, regionprops


from spellchecker import SpellChecker
from TableSeparator import TableSeparator

import openpyxl
import re

from Optimizer import Optimizer

# wb = openpyxl.load_workbook('Files/main.xlsx', data_only=True)
# optimizer = Optimizer(wb)
# optimizer.a()



test = TableSeparator('Files/main.xlsx')
test.get_tables_borders()
test.separate_tables_to_different_worksheets()

