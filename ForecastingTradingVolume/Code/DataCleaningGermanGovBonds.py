#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jan  7 14:29:01 2023

@author: Robert_Hennings
"""

import glob
#load the single data files for storing each Bond in a sheet in an excel file
glob.os.listdir("/Users/Robert_Hennings/Downloads/TradingVolumeForecasting")[5].split(" ")[1]

bonds = []
#create a list of the available bonds
for file in glob.os.listdir("/Users/Robert_Hennings/Downloads/TradingVolumeForecasting"):
    if ".csv" in file:
        bonds.append(file)
        
    else:
        pass
    
#create an excel file that houses all the data
bonds 
book = xw.Book()
try:
    for bond in bonds:
        
        book.sheets.add(bond.split(" ")[1]) 
        
except:
    print(bond)

book.sheets["Tabelle1"].delete()
book.save("/Users/Robert_Hennings/Downloads/TradingVolumeForecasting/Bond_Data.xlsx")
book.close()

#save the data in each sheet
for bond in bonds:
    b = pd.read_csv("/Users/Robert_Hennings/Downloads/TradingVolumeForecasting"+"//"+bond)
    b.columns = ["Date", "Last", "Open", "High", "Low", "PctChange"]
    b.Date = pd.to_datetime(b.Date)
    book = xw.Book("/Users/Robert_Hennings/Downloads/TradingVolumeForecasting/Bond_Data.xlsx")
    book.sheets[bond.split(" ")[1]].range("A1").value = b
    book.save()
    book.close()