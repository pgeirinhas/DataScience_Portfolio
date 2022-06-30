#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Apr 13 16:51:44 2022

@author: pedro.geirinhas
"""

# =============================================================================
# Load Packages and Data
# =============================================================================
# pip install tabula-py
import pandas as pd
from tabula import read_pdf

# Export pdf to a list of Dataframes
raw = tabula.read_pdf("./raw_pdf_data.pdf", pages="all")

# Wrapper function
def df_wrapper(empty_dflist, raw_dflist, correct_cols):
    for df in raw_dflist:
        replace_row = pd.DataFrame(df.columns).T # save value wrongly inserted in title as a new row
        replace_row.columns = correct_cols # correctly rename the columns of the new row dataframe
        df.columns = correct_cols # correctly rename the columns of the dataframe
        df = df.append(replace_row, ignore_index=False, sort=False) # append new row to be part of data
        empty_dflist.append(df) #save it in the newly created dataframe list

#%%# ==========================================================================
# TABLE1 extraction
# Note: Table 1 ends in PDF page 6740
# =============================================================================

# Select Desired Dataframes to Concat divided into 3 sections (Page 1, Page 2 and Page 3)
page1_raw = raw[0:6723:3]
page2_raw = raw[1:6723:3]
page3_raw = raw[2:6723:3]

# Page 1: Append columns that are wrongly read as title and fix title names
correct_cols1 = raw[0].columns
correct_cols2 = raw[1].columns
correct_cols3 = raw[2].columns

# 
page1_clean = []
page2_clean = []
page3_clean = []

# Perform data extraction
df_wrapper(page1_clean, page1_raw, correct_cols1)
df_wrapper(page2_clean, page2_raw, correct_cols2)
df_wrapper(page3_clean, page3_raw, correct_cols3)

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

#%% # =========================================================================
# # TABLE2 extraction
# Note: Table 2 starts in PDF page 6740
# =============================================================================

# Select Desired Dataframes to Concat and append columns that are wrongly read as title
table2_raw = raw[6723:]
correct_cols = raw[6723].columns
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

#%% # Export Concatenated Dataframe to Excel

with pd.ExcelWriter("/Users/pedro.geirinhas/Desktop/glovo.xlsx") as writer:  
    table1_clean.to_excel(writer, sheet_name="table1", index=False)
    table2_clean.to_excel(writer, sheet_name="table2", index=False)