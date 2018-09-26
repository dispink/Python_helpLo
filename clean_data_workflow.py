# -*- coding: utf-8 -*-
"""
Created on Fri Sep 14 12:07:27 2018

@author: Arthur
"""

# 10 minutes to panda: http://pandas.pydata.org/pandas-docs/stable/10min.html#min
# panda cookbook: http://pandas.pydata.org/pandas-docs/stable/cookbook.html#cookbook
# panda cheatsheet: http://pandas.pydata.org/Pandas_Cheat_Sheet.pdf

# import excel
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# assign spreadsheet filename to "file"
file = "MD01_2419-original.xlsx"

# load spreadsheet sec01
xl = pd.ExcelFile(file)

print(xl.sheet_names)

sheet_name = xl.sheet_names[1]


def claen_by_sheet(sheet_name):
    """Itos a function for cleaning data from a given sheet
    """
    # Selecte the desired columnes and exclude the data with NA
    work_sheet = xl.parse(sheet_name, skiprows = 2).dropna(how='any').loc[ : ,
                         ['position (mm)', 'validity','cps', 'MSE', 'Si', 'S', 'Cl', 'Ar', 'K', 'Ca', 'Ti', 'Mn', 'Fe', 'Br', 'Sr']]
    # change the columne name position (mm) to a less bug name position_mm 
    work_sheet.rename(columns={'position (mm)':'position_mm'}, inplace=True)


    # criteria 1: MSE < 1.5 (practicle experience), validity = 1, cps value lower than 1 std (lower limit)
    criteria_1 = (work_sheet.MSE < 1.5) & (work_sheet.validity == 1) & (work_sheet.cps > (work_sheet.cps.mean() - work_sheet.cps.std()))

    work_sheet_1 = work_sheet[criteria_1]

    # criteria 2: Ar/Fe lower than 1 std (upper limit)
    # have to use the data frame pass the criteria 1 since the zero value in Fe makes statics infinite
    ArFe_ratio = work_sheet_1.Ar / work_sheet_1.Fe
    criteria_2 = ArFe_ratio < (ArFe_ratio.mean() + ArFe_ratio.std())

    work_sheet_2 = work_sheet_1[criteria_2]
    
    # criteria 3: exclude the last 2 cm scanning data since they may represent tape instead of core
    bottom_delet = - 20 / (work_sheet_2.position_mm[7] - work_sheet_2.position_mm[6])
    work_sheet_3 = work_sheet_2.loc[0]
    
    return

pd.option_context('display.max_rows', None, 'display.max_columns', None)

with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    print(work_sheet_1)
    print(work_sheet_1[criteria_2].describe())


work_sheet.loc[ : , 'ArFe_ratio'].std()
work_sheet.std()
work_sheet.describe()






