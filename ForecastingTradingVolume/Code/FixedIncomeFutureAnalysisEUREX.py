#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Sep  3 16:55:53 2022

@author: Robert_Hennings
"""

#Load the necessary libraries
import pandas as pd
import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt  
import random
from statsmodels.tsa.arima_model import ARIMA 
import warnings
from  itertools import product
#EUREX Color codes for graphs in HEX notation
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

#Configure the graphic resolution to retina
%config InlineBackend.figure_format = "retina"

#Read in the pure Data on the FI Futures with Price, Open, High, Low and Change in an uncleaned version
#Load the EuroBundFuture as representative example
d = pd.read_excel("/Users/Robert_Hennings/Downloads/FixedIncomeFuturesEurex.xlsx",sheet_name="EuroBundFuture", index_col=0)
#Rename the columns
d.columns = ["Price", "Open", "High", "Low", "Volume", "PctChange"]

#Order the whole dataframe based on the dates
d.sort_values(by="Datum", inplace=True)

#Replace the - with NaNs to drop them afterwards
d.Volume.replace("-", np.nan, inplace=True)
#Drop all NaNs from the Datafrme
d.dropna(inplace=True)
#Add an Index column
d["Ind"] = range(0, d.shape[0])


#Set the tyoe to str to edit the column entries
d.Volume = d.Volume.astype("str")

#If the data is scraped from Investing.com and not downloaded the following cleaning process needs to be executed for each security
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
#Write a small function to plot two variables together
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
#See if there is linear correlation between variables
#Insert the spread as Variable
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
#values for p,d,q should be in the range of values from 0 to 10 for p and q, d in the range of 0 to 2
#create a dataframe with columns p,d,q housing all possible combination sof numbers 0 to 10


d3 = pd.DataFrame(columns=["p","d","q"])
#d is at max 2, values greater than 2 are not supported
d3.p = range(0,11)
d3.d = 2
d3.q = range(0,11)
#use itertools to get all possible combinations

#all columns
d4 = pd.DataFrame(list(product(*d3.values.T)))
d4.drop_duplicates(inplace=True)
d4.columns = ["p","d","q"]
d4.reset_index(drop=True, inplace=True)

#Now the ARIMA model should iterate through each line of the dataframe d4 and be trained
#Results will be saved in a new excel file
book = xw.Book()
book.sheets.add("Summary")
book.sheets["Tabelle1"].delete()
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
    
    
#Read in all the cleaned volume data from each contract, fit the ARIMA(10,2,10) model on it and save the results 

path_read = "/Users/Robert_Hennings/Downloads"
file_name_read = "CleanedFixedIncomeFuturesEUREX.xlsx"

#save the last real values and the last fitted values of the model
#save the predicted 30 values
#save the last 100 values and the predicted 100 values
#save the fitting results
#save everything in one file, each sheet as a product

#Open the existing data containing excel file to read out the sheets 
book = xw.Book(path_read+"//"+file_name_read) 

#save sheets 
sheets = [i.name for i in book.sheets]
#save sheets in new book to keep the structure
new_book = xw.Book()
for k in sheets:
    new_book.sheets.add(k)
    
new_book.sheets["Tabelle1"].delete()
#save the book
new_book.save(path_read+"//"+"ARIMA_Summary_FixedIncome.xlsx")
new_book.close()

#save the needed data into the new file, each sheet represents a product
for table in sheets[0:9]:
    try:
        
        import warnings
        warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARMA', FutureWarning)
        warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARIMA',FutureWarning)
        
        
        book = xw.Book(path_read+"//"+"CleanedFixedIncomeFuturesEUREX.xlsx")
        #Read in the data already existing
        d = pd.DataFrame(book.sheets[table].range("A1").options(expand="table").value)
        #set first row as columns
        d.columns = d.loc[0]
        #drop first column
        d.drop([0], inplace=True)   
        #change datatype of Volume
        d.Volume = d.Volume.astype("float64")
        # fit the model
        arima_model = ARIMA(d.iloc[:,5],order=(10,2,10))
        model = arima_model.fit()
        print("Succes______________")
        b = xw.Book(path_read+"//"+"ARIMA_Summary_FixedIncome.xlsx")
        b.sheets[table].range("A1").value = d.iloc[:,0] #Save the Date as a column
        b.sheets[table].range("B1").value = d.iloc[:,5] #Save the Volume as a column
        #save the fitted in sample data in a separate Dataframe
        fitted = pd.DataFrame(model.fittedvalues[-100:])
        fitted.columns = ["FittedValues"]
        #save 100 model forecasts
        fitted["Predicted"] = model.forecast(100)[0]
        b.sheets[table].range("D1").value = fitted.FittedValues #Last 100 in sample model fitted values
        b.sheets[table].range("F1").value = fitted.Predicted #100 model predicted values
        # b.sheets[table].range("H1").value = model.summary().tables[0].data #Summary of fitted model
        b.save()
        b.close()
    
    except:
        print("Difficulties with: ", table)
    
    
#Now after there is aprediction for every contract with the ARIMA Model, we need to have also a prediction from 
#a multiple linear regression with external variables

#Read in Germany Government Bond data downloaded from investing.com

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

#####################################################################################################################################################
################################################################## Automated Analysis Steps #################################
#What we need: a function that perfroms automatically ARIMA fitting with saving results, plotting and 



def Explorative_ARIMA(data, path_adftable, path_arima_fitting, path_fit_predict_summary, opt_arima_order,num_InSample, num_Predict,data_dates):
    """
    

    Parameters
    ----------
    data : TYPE float, series
        Data onto which an ARIMA Model should be fitted
        Univaraite values of a Timeseries without dates
    path_adftable : TYPE str
        File Path where the ADF Table for Raw Data, First and Second Diff should be saved
    path_arima_fitting : TYPE str
        File Path where the Summary of the ARIMA Fitting process should be saved
    path_fit_predict_summary : TYPE str
        File Path where the results of the final model and its predictions should be saved
    opt_arima_order : TYPE tuple
        Order of the optimal ARIMA Model that generates predictions
    num_InSample : TYPE int
        Number of In Sample Datapoints that should be stored for comparison
    num_Predict : TYPE int
        Number of predictions the fitted ARIMA Model should make
    data_dates : TYPE list, series
        The Dates column from the original Dataframe that houses all the Data

    Returns
    -------
    Three Excel files in the according paths.
    Three plots for the Autocorrelations.

    """
    import warnings
    warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARMA', FutureWarning)
    warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARIMA',FutureWarning)
    import pandas as pd
    import numpy as np
    import xlwings as xw
    import matplotlib.pyplot as plt
    from statsmodels.tsa.arima_model import ARIMA 
    #from statsmodels.tsa.tsatools import adfuller 
    from statsmodels.tsa.stattools import adfuller 
    from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
    #Choosing the parameters p,d,q of the ARIMA Model
    plot_acf(data)
    plt.title("Autocorrelation Raw Data")
    plt.xlabel("Lagged Values")
    plt.ylabel("Correlation")
    plt.show()
    
    #First difference
    plot_acf(data.diff().dropna())
    plt.title("Autocorrelation Fisrt Diff")
    plt.xlabel("Lagged Values")
    plt.ylabel("Correlation")
    plt.show()
    
    #Second difference
    plot_acf(data.diff().diff().dropna())
    plt.title("Autocorrelation Fisrt Diff")
    plt.xlabel("Lagged Values")
    plt.ylabel("Correlation")
    plt.show()
    
    #Test for stationarity with the ADF Test for each level of differenciation
    #Save the results as table
    adftable = pd.DataFrame()
    adftable["Level of Diff"] = ["Raw Data", "First Diff","Second Diff"]
    adftable["P Values"] = [adfuller(data)[1],adfuller(data.diff().dropna())[1],adfuller(data.diff().diff().dropna())[1]]
    #Add Mean and Std and number of data points of the single versions
    adftable["Mean"] = [np.mean(data), np.nanmean(data.diff()), np.nanmean(data.diff().diff())]
    adftable["Std"] = [np.std(data), np.nanstd(data.diff()), np.nanstd(data.diff().diff())]
    adftable["Data Points"] = [len(data), len(data.diff())-1, len(data.diff().diff())-2]
    adftable.set_index(["Level of Diff"])
    print(adftable)
    adfbook = xw.Book()
    adfbook.sheets[0].name = "ADF_Table"
    adfbook.sheets[0].range("A1").options(index=False).value = adftable
    adfbook.save(path_adftable+"//"+"ADF_Table.xlsx")
    adfbook.close()
    
    
    #Find per iteration the best ARIMA Model Fit and store the results in an Excelfile
    #Create all possible values for p,d,q and store in a dataframe thats looped through
    import itertools 
    pdqtable = pd.DataFrame()
    pdqtable["p"] = range(0,11)  
    pdqtable["d"] = 2
    pdqtable["q"] = range(0,11)  
    
    pdqtable_all = pd.DataFrame(list(itertools.product(*pdqtable.values.T)))
    pdqtable_all.drop_duplicates(inplace=True)
    pdqtable_all.reset_index(inplace=True, drop=True)
    
    
    pdqbook = xw.Book()
    pdqbook.sheets.add("Summary")
    pdqbook.sheets["Tabelle1"].delete()
    rowInsheet = 1

    try:
        for p,f,q in zip(pdqtable_all.iloc[:,0], pdqtable_all.iloc[:,1], pdqtable_all.iloc[:,2]):
            warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARMA',FutureWarning)
            warnings.filterwarnings('ignore', 'statsmodels.tsa.arima_model.ARIMA',FutureWarning)
            print(p,f,q)
            arima_model = ARIMA(data, order = (p,f,q))
            model = arima_model.fit()
            # f = open("/Users/Robert_Hennings/Downloads/ARIMAModel/ModelSummary_{p}_{f}_{q}.csv".format(p=p,f=f,q=q), "w")
            #f.write(model.summary().tables[0].data)
            # f.write(model.summary().as_csv())
            # f.close()
            pdqbook.sheets["Summary"].range("A"+str(rowInsheet)).value = model.summary().tables[0].data
            rowInsheet += 8
    except:
        print(p,f,q)
    
    pdqbook.save(path_arima_fitting+"//"+"ARIMA_Optimal_Fitting.xlsx")
    pdqbook.close()
    
        
        
    #Store the AIC for each version p,f,q in a seperate sheet to be able to track the fitting performance
    #still needs to be done here 
    
    #Saving some Data for showing and comparing the model fit
    arima_model = ARIMA(data,order=opt_arima_order)
    model = arima_model.fit()
    
    #open a new book and edit first sheets name
    fit_predict_book = xw.Book()
    fit_predict_book.sheets[0].name = "FitPredict"
    #create Dataframe that stores the desired values
    fitted = pd.DataFrame(model.fittedvalues[-num_InSample:])
    fitted.columns = ["FittedValues "+str(num_InSample)]
    fitted.insert(0,"Date",data_dates[-num_InSample:])
    fitted["Original Values"] = data[-num_InSample:]
    fitted["Predicted "+str(num_InSample)] = model.forecast(num_InSample)[0]
    #dump DataFrame into the excel
    fit_predict_book.sheets[0].range("A1").options(index=False).value = fitted #Last x values in sample model fitted values
    
    
    #save the book at the desired place
    fit_predict_book.save(path_fit_predict_summary+"//"+"Fit_Predict_ModelData.xlsx")
    fit_predict_book.close()
    print("Analysis finished")
    
    
#test the written function
bund = pd.read_excel("/Users/Robert_Hennings/Downloads/CleanedFixedIncomeFuturesEUREX.xlsx", sheet_name="EuroBundFuture")
bund.drop("Ind", axis=1, inplace=True)
bund_dates = bund.Datum 
bund_data = bund.Volume 

#Test the function
Explorative_ARIMA(data=bund_data
                  , path_adftable="/Users/Robert_Hennings/Downloads/test"
                  , path_arima_fitting= "/Users/Robert_Hennings/Downloads/test"
                  , path_fit_predict_summary ="/Users/Robert_Hennings/Downloads/test"
                  , opt_arima_order =(2,2,6)
                  ,num_InSample = 30, 
                  num_Predict=30,
                  data_dates = bund_dates)




#introduce a train test split for the model
def train_test_split(X,y,test_size):
    X_train = X[:round(len(X)*test_size)]
    X_test = X[-round(len(X)*(1-test_size)):]
    y_train = y[:round(len(y)*test_size)]
    y_test = y[-round(len(y)*(1-test_size)):]
    print("Length of X_train: ",len(X_train),
          "\nLength of X_test: ",len(X_test),
          "\nLength of y_train: ",len(y_train),
          "\nLength of y_test: ",len(y_test))
    return X_train, X_test, y_train, y_test
    




f = pd.DataFrame(range(0,20), columns =["Var1"])
f["Target"] = range(60,80)



X_train, X_test, y_train, y_test = train_test_split(X=f.iloc[:,0],y=f.Target,test_size=0.7)

#automatically find the best fit for the residuals 