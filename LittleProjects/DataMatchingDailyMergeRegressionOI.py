# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import glob
import numpy as np

path = "H:\DBAG\FT\900 Work Areas\Robert\DataMatchingDates\\"

glob.os.listdir(path)

for i in glob.os.listdir(path):
    # print(i.split(".")[0])
    # print(path+i)
    globals()[i.split(".")[0]] = pd.read_excel(path+i)
#Merge on the longest available Date Column, that is the one from the First Contract
    
    
SecondContract.columns = ["Date2nd", "SecondOI"]
#Merging all the DataFrames and interpolating missing Data

All = FirstContract.merge(DurationData, how="outer", left_on="Date 1st", right_on="Date").interpolate().merge(SecondContract, how="outer", left_on="Date 1st", right_on="Date2nd").interpolate().merge(ThirdContract, how="outer", left_on="Date 1st", right_on="Date 3rd").interpolate()

All.drop(["Date", "Date2nd", "Date 3rd"], axis=1, inplace=True)


All.head()
All.tail()

d = All.iloc[:,1]
All.pop("1st OI")
All.insert(4,"1st OI",d)

help(pd.DataFrame.insert)



All.set_index("Date 1st", inplace=True)
All.info()
All.to_excel(path+"DataMatchedDates.xlsx")


import sklearn 

dir(skl)

from sklearn import linear_model

dir(linear_model)


help(linear_model.LinearRegression)

RegData = pd.read_excel("H:\DBAG\FT\900 Work Areas\Robert\Variables.xlsx")
RegData.drop("Date ", axis=1, inplace=True)

RegData.dropna(inplace=True)

regression = linear_model.LinearRegression().fit(RegData.iloc[:,1:], RegData.iloc[:,0])

dir(regression)


regression.score(RegData.iloc[:,1:], RegData.iloc[:,0])
regression.coef_