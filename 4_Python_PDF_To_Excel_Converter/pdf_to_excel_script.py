#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Apr 13 16:51:44 2022

@author: pedro.geirinhas
"""

# =============================================================================
# Load Packages and Data
# =============================================================================

import pandas as pd
import tabula

# Export pdf to a list of Dataframes
raw = tabula.read_pdf("/Users/pedro.geirinhas/Downloads/Retail_Project_Manager_Business_Case.pdf", pages="all")

# Wrapper function
def df_wrapper(empty_dflist, raw_dflist, correct_cols):
    for df in raw_dflist:
        replace_row = pd.DataFrame(df.columns).T # save value wrongly inserted in title as a new row
        replace_row.columns = correct_cols # correctly rename the columns of the new row dataframe
        df.columns = correct_cols # correctly rename the columns of the dataframe
        df = df.append(replace_row, ignore_index=False, sort=False) # append new row to be part of data
        empty_dflist.append(df) #save it in the newly created dataframe list


#%% # =============================================================================
# # TABLE2 extraction
# page 6740 ≠ table >> table 2
# =============================================================================

# Select Desired Dataframes to Concat and append columns that are wrongly read as title
table2_raw = raw[6741:]
correct_cols = raw[6740].columns
table2_dflist = []

# Perform data extraction
df_wrapper(table2_dflist, table2_raw, correct_cols)

# Concat clean Dataframes and perform some Dataframe wrangling for format optimization
table2_clean = pd.concat(table2_dflist)
table2_clean = pd.concat([table2_clean, raw[6740]], ignore_index=False, axis=0)
table2_clean.reset_index(drop=True, inplace=True)
table2_clean.ID = pd.to_numeric(table2_clean.ID, errors="coerce")
table2_clean["item cost"] = pd.to_numeric(table2_clean["item cost"], errors="coerce")
table2_clean.price = pd.to_numeric(table2_clean.price, errors="coerce")
table2_clean["Unit sales"] = pd.to_numeric(table2_clean["Unit sales"], errors="coerce")
table2_clean = table2_clean.sort_values("ID", ascending=True, na_position="last")

#%%# =============================================================================
# TABLE1 extraction
# Note: Tabula unable to read page 17 of pdf
# =============================================================================

# Select Desired Dataframes to Concat divided into 3 sections (Page 1, Page 2 and Page 3)
# Need to create list a & b since tabula cannot read page 18 which breaks flow of pdf pages and for loop
page1_lista = raw[3:13:3]
page1_listb = raw[17:6740:3]
page1_list = [page1_lista, page1_listb]

page2_lista = raw[4:14:3]
page2_listb = raw[18:6740:3]
page2_list = [page2_lista, page2_listb]

page3_lista = raw[5:15:3]
page3_listb = raw[19:6740:3]
page3_list = [page3_lista, page3_listb]

# Page 1: Append columns that are wrongly read as title and fix title names
correct_cols = raw[0].columns
page1_dflist = []

# Perform data extraction
for df_list in page1_list:
    df_wrapper(page1_dflist, df_list, correct_cols)

#Merge both Dataframe lists into a single Page1 Dataframe
page1_clean = pd.concat(page1_dflist)
page1_clean = pd.concat([page1_clean, raw[0]], ignore_index=False, axis=0)

# Page 2: Append columns that are wrongly read as title and fix title names
correct_cols = raw[1].columns
page2_dflist = []

# Perform data extraction
for df_list in page2_list:
    df_wrapper(page2_dflist, df_list, correct_cols)

#Merge both Dataframe lists into a single Page2 Dataframe
page2_clean = pd.concat(page2_dflist)
page2_clean = pd.concat([page2_clean, raw[1]], ignore_index=False, axis=0)

# Page 3: Append columns that are wrongly read as title and fix title names
correct_cols = raw[2].columns
page3_dflist = []

# Perform data extraction
for df_list in page3_list:
    df_wrapper(page3_dflist, df_list, correct_cols)

#Merge both Dataframe lists into a single Page3 Dataframe
page3_clean = pd.concat(page3_dflist)
page3_clean = pd.concat([page3_clean, raw[2]], ignore_index=False, axis=0)

# Concatenate 3 Dataframes into 1 and perform some Dataframe wrangling for format optimization
table1_clean = pd.concat([page1_clean, page2_clean, page3_clean],ignore_index=False, axis=1)
table1_clean["PRODUCT ID"] = pd.to_numeric(table1_clean["PRODUCT ID"], errors="coerce")
table1_clean["EXTERNAL ID"] = pd.to_numeric(table1_clean["EXTERNAL ID"], errors="coerce")
table1_clean["PRODUCT PRICE"] = pd.to_numeric(table1_clean["PRODUCT PRICE"], errors="coerce")
table1_clean["QUANTITY"] = pd.to_numeric(table1_clean["QUANTITY"], errors="coerce")
table1_clean["ORDER ID"] = pd.to_numeric(table1_clean["ORDER ID"], errors="coerce")
table1_clean.reset_index(drop=True, inplace=True)
table1_clean = table1_clean.convert_dtypes()
table1_clean["CREATION LOCAL TIME"] = pd.to_datetime(table1_clean["CREATION LOCAL TIME"], errors="coerce")

#%% # Export Concatenated Dataframe to Excel

with pd.ExcelWriter("/Users/pedro.geirinhas/Desktop/glovo.xlsx") as writer:  
    table1_clean.to_excel(writer, sheet_name="table1", index=False)
    table2_clean.to_excel(writer, sheet_name="table2", index=False)