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
    data_ex_map = pd.DataFrame()
    # criteria 1: exclude the last 2 cm scanning data since they may represent tape instead of core
    resolution = work_sheet.position_mm[7] - work_sheet.position_mm[6]
    bottom_delet = int(20 / resolution)
    # the first position is an overlapped measurement to the last position in previous section when it's 0
    if work_sheet.position_mm.iloc[0] == 0:
        work_sheet_1 = work_sheet[1: -bottom_delet].copy()
    else:
        work_sheet_1 = work_sheet[ : -bottom_delet].copy()
    data_ex_1 = work_sheet[~work_sheet.position_mm.isin(work_sheet_1.position_mm)].copy()
    data_ex_1['failed_at_criteria'] = ['1' for _ in data_ex_1.position_mm]

    # criteria 2: validity = 1, cps value lower than 3 std (lower limit), Fe > 0
    criteria_2 = (
            (work_sheet_1.validity == 1) 
            & (work_sheet_1.cps > (work_sheet.cps.mean() - 3 * work_sheet_1.cps.std()))
            & (work_sheet_1.Fe > 0)
            )
    work_sheet_2 = work_sheet_1[criteria_2].copy()
    data_ex_2 = work_sheet_1[~work_sheet_1.position_mm.isin(work_sheet_2.position_mm)].copy()
    data_ex_2['failed_at_criteria'] = ['2' for _ in data_ex_2.position_mm]
  
    # criteria 3: Ar/Fe lower than 3 std (upper limit)
    ArFe_ratio = work_sheet_2.Ar / work_sheet_2.Fe
    criteria_3 = ArFe_ratio < (ArFe_ratio.mean() + 3 * ArFe_ratio.std())
    work_sheet_3 = work_sheet_2[criteria_3].copy()
    data_ex_3 = work_sheet_2[~work_sheet_2.position_mm.isin(work_sheet_3.position_mm)].copy()
    data_ex_3['failed_at_criteria'] = ['3' for _ in data_ex_3.position_mm]
    
    # label the section name in this dataframe
    work_sheet_3['section'] = [sheet_name for i in work_sheet_3.Fe]
    
    # output the cleaned data and excluded data
    data_ex_map = data_ex_1.append(data_ex_2).append(data_ex_3)
    excluded_percentage = len(data_ex_map)/len(work_sheet) * 100
    writer = pd.ExcelWriter('cleaned_MD01_2419_{}_{}%.xlsx'.format(sheet_name, round(excluded_percentage)))
    work_sheet_3.drop('section', axis = 1).to_excel(writer,'cleaned_data', index = False)
    data_ex_map.to_excel(writer,'excluded_data', index = False)
    writer.save()
    
    return work_sheet_3


# Run all the sections 
cleaned_data_map = pd.DataFrame()
sec_count = 0
end_position_list = [0]     # give a list end_position_list and give 0 to the first value
rep = 1
for sheet_name in xl.sheet_names[1: len(xl.sheet_names)]:
    if len(sheet_name) == 8:        # sections having two replicates go into this loop
        if rep == 2:                    # sec**_r2 go into this loop
            sec_count += 1
            cleaned_data_2 = clean_by_sheet(sheet_name)
            data_merged = pd.merge(cleaned_data_1,
                                   cleaned_data_2,
                                   how = 'inner',
                                   on = 'position_mm'
                                   )
            mean_data = pd.DataFrame()
            mean_data['position_mm'], mean_data['validity'] = data_merged.position_mm, data_merged.validity_x
            # make means of each value between two replicates
            for column in cleaned_data_1.columns[2 : -1]:
                mean_data['{}'.format(column)] = (
                        (data_merged['{}_x'.format(column)] + data_merged['{}_y'.format(column)]) / 2
                        )
            mean_data['section'] = ['{:.5}'.format(sheet_name) for _ in mean_data.position_mm]
            # take only end_position from r2. use the end posistion from last section to calibrate position
            mean_data['cal_position_mm'] = mean_data.position_mm + end_position_list[sec_count-1]
            end_position_list.append(mean_data.cal_position_mm.iloc[-1])
            cleaned_data_map = cleaned_data_map.append(mean_data)
            rep = 1            
        else:                           # sec**_r2 go into this loop
            cleaned_data_1= clean_by_sheet(sheet_name)
            rep += 1
    else:                           # sections don't having replicates go into this loop
        sec_count += 1
        cleaned_data = clean_by_sheet(sheet_name)
        cleaned_data['cal_position_mm'] = cleaned_data.position_mm + end_position_list[sec_count-1]
        end_position_list.append(cleaned_data.cal_position_mm.iloc[-1])
        cleaned_data_map = cleaned_data_map.append(cleaned_data)

# Final: output the compiled data in csv        
cleaned_data_map.to_csv('MD01_2419_all_sections.csv', index = False)