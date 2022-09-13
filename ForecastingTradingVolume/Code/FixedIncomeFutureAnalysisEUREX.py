#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Sep  3 16:55:53 2022

@author: Robert_Hennings
"""

import pandas as pd
import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt  
import random
from statsmodels.tsa.arima_model import ARIMA 
import warnings
#EUREX Color codes
#9ebedf
#3c7dbe
#80e7be
#00ce7d
#908ba8
#201751
#808080
#E7E6E6
#7030A0
#A5A5A5

#Config the graphic resolution
%config InlineBackend.figure_format = "retina"

#Read in the pure Data on the FI Futures with Price, Open, High, Low and Change
d = pd.read_excel("/Users/Robert_Hennings/Downloads/FixedIncomeFuturesEurex.xlsx",sheet_name="EuroBundFuture", index_col=0)
#Edit the columns
d.columns = ["Price", "Open", "High", "Low", "Volume", "PctChange"]

#Order the whole dataframe in a new way
d.sort_values(by="Datum", inplace=True)

#Replace the - with NaNs
d.Volume.replace("-", np.nan, inplace=True)
#Drop all NaNs 
d.dropna(inplace=True)
#Add an Index column
d["Ind"] = range(0, d.shape[0])


#Set the tyoe to str to edit the column entries
d.Volume = d.Volume.astype("str")


#Replace every M and multiply the entry by 1000000 and do the same but with thousands for the entries containing K
for i in d.Volume:
    if "M" in str(i):
          d.iloc[d.Ind[d.Volume ==i],4] = float(i.split("M")[0].replace(",","."))*1000000
        
    elif "K" in str(i):
        d.iloc[d.Ind[d.Volume ==i],4] = float(i.split("K")[0].replace(",","."))*1000
    
#Convert Volume to Thousands and reset the column type to float
d.Volume = d.Volume.astype("float64")
d.Volume = d.Volume/1000 
    
    
#Compute the Volatility of the Volume in two ways
d["Std_Volume"] = ""

#Value for the Volatility on each day, expanding the respected area by one day
for i in range(0,d.shape[0]):
    d.iloc[i,7] = d.iloc[0:i,4].std()    

#Compute Std with a rolling 30 day window
d["MA30_Std_Volume"] = d.Volume.rolling(30).std()
    


#Compute the volatility of the Price 
d["Std_Price"] = ""

for i in range(0,d.shape[0]):
    d.iloc[i,9] = d.iloc[0:i,0].std() 



#Plot the Daily Volume and its Volatility
#Write a small function to plot to variables together
def plot_days(data,column_names,colors,labels,days, title,ylabel,xlabel):
    plt.plot(data[column_names[0]][-days:], color=colors[0], label=labels[0])
    plt.plot(data[column_names[1]][-days:], color=colors[1], label=labels[1])
    plt.legend()
    plt.title(title+str(days)+" days")
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.xticks(rotation=45)
    plt.show()
    

#last 30 days
plot_days(data=d,column_names=["Volume", "Std_Volume"],colors=["#201751","#00ce7d"],labels=["Volume", "Volatility Volume"],
          days=30, title="Volume and its Volatility ",ylabel="Price and Volume",xlabel="Time")
#all days
plot_days(data=d,column_names=["Volume", "Std_Volume"],colors=["#201751","#00ce7d"],labels=["Volume", "Volatility Volume"],
          days=d.shape[0], title="Volume and its Volatility ",ylabel="Price and Volume",xlabel="Time")




#Compute the Volatility of the returns of the price 
d["Std_PctChange"] = ""

for i in range(0,d.shape[0]):
    d.iloc[i,10] = d.iloc[0:i,6].std() 



#Compute the Vola on a weekly level
d["Std_PctChangeWeekly"] = d.PctChange.rolling(7).std()



#plot the last 30 days, Volume and Std_PctChange
plot_days(data=d,column_names=["Volume", "Std_PctChange"],colors=["#201751","#00ce7d"],labels=["Volume", "Volatility Returns"],
          days=30, title="Volume and Return Vola ",ylabel="Price and Vola Returns",xlabel="Time")
#all days Volume and Std_PctChange
plot_days(data=d,column_names=["Volume", "Std_PctChange"],colors=["#201751","#00ce7d"],labels=["Volume", "Volatility Returns"],
          days=d.shape[0], title="Volume and Return Vola ",ylabel="Price and Vola Returns",xlabel="Time")

#plot the last 30 days, Volume and Std_PctChangeWeekly
plot_days(data=d,column_names=["Volume", "Std_PctChangeWeekly"],colors=["#201751","#00ce7d"],labels=["Volume", "Volatility Returns"],
          days=30, title="Volume and Return Vola Weekly ",ylabel="Price and Vola Returns",xlabel="Time")


#Now investigate to find a proper ARIMA Model to forecast the volume
#See if theres linear correlation between variables

d.insert(6,"SpreadHigh_Low",value=d.High-d.Low)

#As some papers state, there is a certain linear correlation between the High-Low Spread and the Trading Volume
d.iloc[:,0:7].corr()

#Employ a simple linear regression in the current state of the project
from sklearn.linear_model import LinearRegression

y = d.Volume
d2 = d.copy()
d2.drop(["Volume"],axis=1, inplace=True)
model = LinearRegression()
model.fit(X=d2.iloc[:,0:5],y=y)
dir(model)
model.score(X=d2.iloc[:,0:5],y=y)




d.Volume.plot(kind="kde")
plt.show()


np.log(d.Volume).plot(kind="kde")
plt.show()





#Create an iterating algorithm that iterates through all combinations for p,d,q variables in an ARIMA model
#values for p,d,q should be n the range of 10
#create a dataframe with columns p,d,q housing all possible combination sof numbers 0 to 10


d3 = pd.DataFrame(columns=["p","d","q"])
#d is at max 2, values greater than 2 are not supported
d3.p = range(0,11)
d3.d = 2
d3.q = range(0,11)

from  itertools import product
#all columns
d4 = pd.DataFrame(list(product(*d3.values.T)))
d4.drop_duplicates(inplace=True)
d4.columns = ["p","d","q"]
d4.reset_index(drop=True, inplace=True)

#Now the ARIMA model should iterate through each line of the dataframe d4

book = xw.Book()
book.sheets.add("Summary")
rowInsheet = 1

try:
    for p,f,q in zip(d4.iloc[24:,0], d4.iloc[24:,1], d4.iloc[24:,2]):
        warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARMA',FutureWarning)
        warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARIMA',FutureWarning)
        print(p,f,q)
        arima_model = ARIMA(d.iloc[:,4].values, order = (p,f,q))
        model = arima_model.fit()
        # f = open("/Users/Robert_Hennings/Downloads/ARIMAModel/ModelSummary_{p}_{f}_{q}.csv".format(p=p,f=f,q=q), "w")
        #f.write(model.summary().tables[0].data)
        # f.write(model.summary().as_csv())
        # f.close()
        book.sheets["Summary"].range("A"+str(rowInsheet)).value = model.summary().tables[0].data
        rowInsheet += 8
except:
    print(p,f,q)
    
    
#Read in all the cleaned volume data from each contract, fit the ARIMA(10,2,10) model on it and save some results 

path_read = "/Users/Robert_Hennings/Downloads"
file_name_read = "CleanedFixedIncomeFuturesEUREX.xlsx"

#save the last real values and the last fitted values of the model
#save the predicted 30 values
#save the last 100 values and the predicted 100 values
#save the fitting results
#save everything in one file, each sheet as a product

book = xw.Book(path_read+"//"+file_name_read) 

sheets = [i.name for i in book.sheets]
new_book = xw.Book()
for k in sheets:
    new_book.sheets.add(k)
    
new_book.sheets["Tabelle1"].delete()

new_book.save(path_read+"//"+"ARIMA_Summary_FixedIncome.xlsx")
new_book.close()

#save the needed data into the new file, each sheet represents a product
for table in sheets[0:9]:
    try:
        
        import warnings
        warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARMA', FutureWarning)
        warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARIMA',FutureWarning)
        
        
        book = xw.Book(path_read+"//"+"CleanedFixedIncomeFuturesEUREX.xlsx")
        d = pd.DataFrame(book.sheets[table].range("A1").options(expand="table").value)
        d.columns = d.loc[0]
        d.drop([0], inplace=True)   
        d.Volume = d.Volume.astype("float64")
        # print(d.head())
        # arima_model = ARIMA(d.iloc[:,5],order=(10,2,10))
        # model = arima_model.fit()
        print("Succes______________")
        b = xw.Book(path_read+"//"+"ARIMA_Summary_FixedIncome.xlsx")
        b.sheets[table].range("A1").value = d.iloc[:,0] #Save the Date as a column
        b.sheets[table].range("B1").value = d.iloc[:,5] #Save the Volume as a column
        
        fitted = pd.DataFrame(model.fittedvalues[-100:])
        fitted.columns = ["FittedValues"]
        fitted["Predicted"] = model.forecast(100)[0]
        b.sheets[table].range("D1").value = fitted.FittedValues #Last 100 in sample model fitted values
        b.sheets[table].range("F1").value = fitted.Predicted #100 model predicted values
        # b.sheets[table].range("H1").value = model.summary().tables[0].data #Summary of fitted model
        b.save()
        b.close()
    
    except:
        print("Difficulties with: ", table)
    
    
    





