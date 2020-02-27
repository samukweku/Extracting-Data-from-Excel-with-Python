#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 21:21:04 2020

@author: sam
"""

from pathlib import Path
from fnmatch import fnmatch
import pandas as pd

folder = Path('004-MSPTDA-ExcelFiles')

all_files = list(folder.iterdir())

excel_only_files = [x for x in folder.iterdir()
                    if fnmatch(x.name.lower(),'*.xls*')
                   ]


def process_files(xls_file):
    '''
    Picks an excel file,
    extracts the city name using the stem method from Path,
    filters out any sheet name that starts with Sheet,
    reads in the excel file using pandas, with the specified sheet names,
    concatenates the resulting dictionary,
    assigns city back to the dataframe
    and returns a dataframe
    '''
    
    xls = pd.ExcelFile(xls_file)
    
    city = xls_file.stem
    
    sheet_list = [sheet for sheet in xls.sheet_names                  
                  if not sheet.startswith('Sheet')]
    
    df = pd.read_excel(xls,sheet_list)
    
    outcome = pd.concat([data.assign(SalesPerson = SalesPerson) 
                         for SalesPerson, data in df.items()],
                        ignore_index=True
                       ).assign(City = city)
    
    return outcome    

combo = [process_files(xls_file) for xls_file in excel_only_files]

outcome = pd.concat(combo, ignore_index=True)