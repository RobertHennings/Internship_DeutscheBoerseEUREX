#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep  9 16:06:34 2022

@author: Robert_Hennings
"""



import pandas as pd
import numpy as np
import xlwings as xw
import warnings
import statsmodels 
from statsmodels.tsa import arima_model 


#First set upm a cleaned version of all the Fixed Income data 
#remove Ms and Ks and multiply accordingly, save in a new file in a new sheet 
path_read = "/Users/Robert_Hennings/Downloads"
file_name_read = "FixedIncomeFuturesEUREX.xlsx"

book = xw.Book(path_read+"//"+file_name_read)

sheets = [i.name for i in book.sheets] 

sheets.sort(reverse=True)
sheets.pop(0)
sheets.pop(11)

new_book = xw.Book()

for j in sheets:
    new_book.sheets.add(j)
    
    
new_book.sheets["Tabelle1"].delete()
new_book.save(path_read+"//"+"CleanedFixedIncomeFuturesEUREX.xlsx")
new_book.close()

#use only the futures sheets beacuse i havent added the data for the options, that would generate errors
for i in sheets[11:]:
    try: 
     globals()[f"{i}_df"] = pd.read_excel(path_read+"//"+file_name_read, sheet_name=i, index_col=0)
    
     globals()[f"{i}_df"].columns = ["Price", "Open", "High", "Low", "Volume", "PctChange"]

    #Order the whole dataframe in a new way
     globals()[f"{i}_df"].sort_values(by="Datum", inplace=True)
    
    #Replace the - with NaNs
     globals()[f"{i}_df"].Volume.replace("-", np.nan, inplace=True)
    #Drop all NaNs 
     globals()[f"{i}_df"].dropna(inplace=True)
    #Add an Index column
     globals()[f"{i}_df"]["Ind"] = range(0, globals()[f"{i}_df"].shape[0])
    
     #Set the tyoe to str to edit the column entries
     globals()[f"{i}_df"].Volume = globals()[f"{i}_df"].Volume.astype("str")
    
    
     #Replace every M and multiply the entry by 1000000 and do the same but with thousands for the entries containing K
     for v in globals()[f"{i}_df"].Volume:
        if "M" in str(v):
              globals()[f"{i}_df"].iloc[globals()[f"{i}_df"].Ind[globals()[f"{i}_df"].Volume ==v],4] = float(v.split("M")[0].replace(",","."))*1000000
            
        elif "K" in str(v):
            globals()[f"{i}_df"].iloc[globals()[f"{i}_df"].Ind[globals()[f"{i}_df"].Volume ==v],4] = float(v.split("K")[0].replace(",","."))*1000
        
     #Convert Volume to Thousands and reset the column type to float
     globals()[f"{i}_df"].Volume = globals()[f"{i}_df"].Volume.astype("float64")
     book = xw.Book(path_read+"//"+"CleanedFixedIncomeFuturesEUREX.xlsx")
     # book.sheets.add(i)
     book.sheets[i].range("A1").value = globals()[f"{i}_df"]
     book.save()
     book.close()
     
    except:
        print("Difficulties with: ", i)
    
    
    
    
    






