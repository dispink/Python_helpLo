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
#import numpy as np
#import matplotlib.pyplot as plt

# assign spreadsheet filename to "file"
file = "MD01_2419-original.xlsx"

# load spreadsheet sec01
xl = pd.ExcelFile(file)


def clean_by_sheet(sheet_name):
    """It's a function for cleaning data from a given sheet.
    It automatically outputs a xlsx file containing cleaned data and excluded data.
    Then, it returns the dataframe, end position of this section, percentage of the excluded data.
    """
    # Selecte the desired columnes and exclude the data with NA
    # then change the columne name position (mm) to a less bug name position_mm 
    work_sheet = (xl.parse(sheet_name, skiprows = 2).dropna(how = 'any').loc[ : ,
                         ['position (mm)', 'validity','cps', 'MSE', 'Si', 'S', 'Cl',
                          'Ar', 'K', 'Ca', 'Ti', 'Mn', 'Fe', 'Br', 'Sr']]
                         .rename(columns = {'position (mm)' : 'position_mm'})
                         )

    
    # criteria 1: exclude the last 2 cm scanning data since they may represent tape instead of core
    bottom_delet = int(20 / (work_sheet.position_mm[7] - work_sheet.position_mm[6]))
    work_sheet_1 = work_sheet[ : -bottom_delet]

    # criteria 2: MSE value lower than 2 std (higher limit), validity = 1, cps value lower than 2 std (lower limit)
    criteria_2 = (
            (work_sheet_1.MSE < (work_sheet_1.MSE.mean() + 2 * work_sheet_1.MSE.std())) 
            & (work_sheet_1.validity == 1) 
            & (work_sheet_1.cps > (work_sheet.cps.mean() - 2 * work_sheet_1.cps.std()))
            )
    work_sheet_2 = work_sheet_1[criteria_2]

    # criteria 3: Ar/Fe lower than 1 std (upper limit)
    # have to use the data frame pass the criteria 2 since the zero value in Fe makes statics infinite
    ArFe_ratio = work_sheet_2.Ar / work_sheet_2.Fe
    criteria_3 = ArFe_ratio < (ArFe_ratio.mean() + ArFe_ratio.std())
    work_sheet_3 = work_sheet_2[criteria_3]
    

    # label the section name in this dataframe
    work_sheet_3['section'] = [sheet_name for i in work_sheet_3.Fe]
    
    # mark the last valid position of each core section
    end_position = work_sheet_3.iloc[-1, 0]
    
    # output the cleaned data and excluded data
    work_sheet_ex = work_sheet[~work_sheet.position_mm.isin(work_sheet_3.position_mm)]
    excluded_percentage = len(work_sheet_ex)/len(work_sheet_3) * 100
    writer = pd.ExcelWriter('cleaned_MD01_2419_{}_{}%.xlsx'.format(sheet_name, round(excluded_percentage)))
    work_sheet_3.drop('section', axis = 1).to_excel(writer,'cleaned_data', index = False)
    work_sheet_ex.to_excel(writer,'excluded_data', index = False)
    writer.save()
    
    return work_sheet_3, end_position

print(xl.sheet_names)

# deal with the sections have 2 replicates ####still building....
cleaned_data_map = pd.DataFrame()
sec_count = 0
end_position_list = [0]     # give a list end_position_list and give 0 to the first value
rep = 1
for sheet_name in xl.sheet_names[1:4]:
    if len(sheet_name) == 8:        # sections having two replicates go into this loop
        if rep == 2:                    # sec**_r2 go into this loop
            cleaned_data_2, end_position = clean_by_sheet(sheet_name)
            data_merged = pd.merge(cleaned_data_1,
                                   cleaned_data_2,
                                   how = 'inner',
                                   on = 'position_mm'
                                   )
            mean_data = pd.DataFrame()
            mean_data['position_mm'], mean_data['validity'] = data_merged.position_mm, data_merged.validity_x
            for column in cleaned_data_1.columns[2 : -1]:
                mean_data['{}'.format(column)] = (
                        (data_merged['{}_x'.format(column)] + data_merged['{}_y'.format(column)]) / 2
                        )
            mean_data['section'] = ['{:.5}'.format(sheet_name) for _ in mean_data.position_mm]
            cleaned_data_map = cleaned_data_map.append(mean_data)
            # take only end_position from r2
            end_position_list.append(end_position)
            mean_data['cal_postion_mm'] = mean_data.position_mm
            rep = 1
            sec_count += 1
        else:                           # sec**_r2 go into this loop
            cleaned_data_1, end_position = clean_by_sheet(sheet_name)
            rep += 1
    else:                           # sections don't having replicates go into this loop
        cleaned_data, end_position = clean_by_sheet(sheet_name)
        cleaned_data_map = cleaned_data_map.append(cleaned_data)
        end_position_list.append(end_position)
        sec_count += 1




end_position_list = []
for sheet_name in xl.sheet_names[1:3]:
    cleaned_data, end_position, excluded_percentage = clean_by_sheet(sheet_name)
    cleaned_data_map['{}'.format(sheet_name)] = cleaned_data
    end_position_list.append(end_position)

with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    print(work_sheet_1)
    print(work_sheet_1[criteria_2].describe())








