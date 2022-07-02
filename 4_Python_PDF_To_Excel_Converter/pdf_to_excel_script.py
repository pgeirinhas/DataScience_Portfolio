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
from tabula import read_pdf # Note: pip install tabula-py

# Export pdf with the Tabula Package that generates a List of Dataframe (each page of the PDF is a Dataframe) 
raw = read_pdf("./raw_pdf_data.pdf", pages="all")

#%%# ==========================================================================
# Define User-Defined Functions for Script
# =============================================================================

# Wrapper function that cleans wrongly interpreted data by Tabula
def df_wrapper(raw_dflist: list, correct_colnames) -> pd.DataFrame():
    """Returns a clean Dataframe that will be Exported to Excel as a table.
    
    Keyword arguments:
    raw_dflist -- a list of Pandas Data Frames to be corrected 
    correct_colnames -- the desired column names (Can be a list of strings or Pandas DataFrame Index)
    """
    clean_df = pd.DataFrame(columns=correct_colnames) # empty Dataframe

    for df in raw_dflist:
        misplaced_row = df.columns.to_list() # save values row that Tabula wrongly interpets as column names of each Dataframe
        df.loc[len(df),:] = misplaced_row # append wrongly interpreted row by Tabula in correct place of each Dataframe
        df.columns = correct_colnames # correctly rename the columns of each Dataframe that Tabula wrongly interpreted
        clean_df = pd.concat([clean_df, df], axis=0) # generate a single Dataframe with all corrected data    
    
    return clean_df

#%%# ==========================================================================
# Find What Page TABLE2 Starts Is Located In PDF 
# =============================================================================

pdf_tables = [raw[0].shape,raw[1].shape,raw[2].shape] # by inspecting the PDF the first 3 pages consist of table 1 
pdf_table2 = [] 
for df,row in zip(raw, range(0,len(raw)-1)):
    if df.shape in pdf_tables or df.shape[1] in [2,4]: # by inspecting the PDF the first 3 pages have shape (,2) and (,4)
        continue
    else:
        pdf_tables.append(df.shape) # update pdf_tables list when new table is found 
        pdf_table2.append(tuple((df.shape[0], df.shape[1],row))) 

table2_pdfpage = pdf_table2[0][2]

#%%# ==========================================================================
# TABLE1 extraction
# Note: Table 1 ends in PDF page 6726
# =============================================================================

# Select TABLE1 raw data & column names that were wrongly interpreted by Tabula -- divide them into 3 sections (Page 1, Page 2 and Page 3)
page1_raw = raw[0:table2_pdfpage:3]
correct_cols1 = raw[0].columns

page2_raw = raw[1:table2_pdfpage:3]
correct_cols2 = raw[1].columns

page3_raw = raw[2:table2_pdfpage:3]
correct_cols3 = raw[2].columns

# Perform data cleaning & concatenate horizontally 3 sections into 1
page1_clean = df_wrapper(page1_raw, correct_cols1)
page2_clean = df_wrapper(page2_raw, correct_cols2)
page3_clean = df_wrapper(page3_raw, correct_cols3)
table1_clean = pd.concat([page1_clean, page2_clean, page3_clean],ignore_index=False, axis=1)

# Perform data wrangling for format optimization
table1_clean["PRODUCT ID"] = pd.to_numeric(table1_clean["PRODUCT ID"], errors="coerce")
table1_clean["EXTERNAL ID"] = pd.to_numeric(table1_clean["EXTERNAL ID"], errors="coerce")
table1_clean["PRODUCT PRICE"] = pd.to_numeric(table1_clean["PRODUCT PRICE"], errors="coerce")
table1_clean["QUANTITY"] = pd.to_numeric(table1_clean["QUANTITY"], errors="coerce")
table1_clean["ORDER ID"] = pd.to_numeric(table1_clean["ORDER ID"], errors="coerce")
table1_clean.reset_index(drop=True, inplace=True)
table1_clean = table1_clean.convert_dtypes()
table1_clean["CREATION LOCAL TIME"] = pd.to_datetime(table1_clean["CREATION LOCAL TIME"], errors="coerce")

#%%# ==========================================================================
# # TABLE2 extraction
# Note: Table 2 starts in PDF page 6727
# =============================================================================

# Select TABLE2 raw date & column names that were wrongly interpreted by Tabula
table2_raw = raw[table2_pdfpage:]
correct_cols = raw[table2_pdfpage].columns

# Perform data cleaning & extraction
table2_clean = df_wrapper(table2_raw, correct_cols)

# Perform data wrangling for format optimization
table2_clean.reset_index(drop=True, inplace=True)
table2_clean.ID = pd.to_numeric(table2_clean.ID, errors="coerce")
table2_clean["item cost"] = pd.to_numeric(table2_clean["item cost"], errors="coerce")
table2_clean.price = pd.to_numeric(table2_clean.price, errors="coerce")
table2_clean["Unit sales"] = pd.to_numeric(table2_clean["Unit sales"], errors="coerce")
table2_clean = table2_clean.sort_values("ID", ascending=True, na_position="last")

#%% # Export Concatenated Dataframes as tables to Excel in seperate sheets 

with pd.ExcelWriter("./clean_excel.xlsx") as writer:  
    table1_clean.to_excel(writer, sheet_name="table1", index=False)
    table2_clean.to_excel(writer, sheet_name="table2", index=False)