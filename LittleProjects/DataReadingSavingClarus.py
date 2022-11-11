# -*- coding: utf-8 -*-
"""
Created on Tue Aug  2 13:50:10 2022

@author: ka710
"""
import pandas as pd

path = "S:\A_KREISE\ZINSEN\Reporting"+"\\"+"2022"+"\\" +"HY1 Report\STIR\Charts\ClarusDataUpdateRH"

df = pd.read_csv(path+"\\"+"TotalIRD_DV01_TradedPerMonthRawData.csv")

df.to_excel(path+"\\"+"TotalIRD_DV01_TradedPerMonthExcelData.xlsx", index=False)
