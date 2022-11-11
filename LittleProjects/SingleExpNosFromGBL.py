# -*- coding: utf-8 -*-
"""
Created on Thu Aug  4 15:12:49 2022

@author: ka710
"""
import pandas as pd

#Set the path to the folder that holds the text data that should be read in
path = "H:\DBAG\FT\900 Work Areas\Robert\\"+"04"+"RegressionVolumesRevenues\RegressionOpenInterest\Data" 
#provide the filename
file_name = "Raw_Data_OpenIntperProduct.txt"

#Read in the text file that holds the OI for the FGBL contract(s)
OIdata = pd.read_table(path+"\\"+file_name, header=0)
#As the column names are the first row of the data, there has to be inserted a new first row
OIdata = OIdata.shift(1)
#take the values from the column names and store them in the newl created first row
OIdata.loc[0] = OIdata.columns
#set new column names
OIdata.columns = ["Date", "ExpNo", "Product", "OI"]

#Separate the whole file into the single Exp No that range from 1 to 3
FGBL_1 = OIdata[OIdata.ExpNo == 1]

FGBL_2 = OIdata[OIdata.ExpNo == 2]

FGBL_3 = OIdata[OIdata.ExpNo == 3]
#Compare the lenghts of the different Contract data
print("Lenghth of different Exp No for the FGBL: ", FGBL_1.shape[0],FGBL_2.shape[0], FGBL_3.shape[0])
#Reset the index so that it will start for every frame again at 0 as usually
FGBL_1.reset_index(inplace=True, drop=True)
FGBL_2.reset_index(inplace=True, drop=True)
FGBL_3.reset_index(inplace=True, drop=True)

#save the single Exp Nos in single Excel files in the provided path 
for i in range(1,4):
    # print("FGBL_"+str(i))
    globals()["FGBL_"+str(i)].to_excel(path+"\\"+"FGBL_ExpNo"+str(i)+"_OI.xlsx", index=False)

