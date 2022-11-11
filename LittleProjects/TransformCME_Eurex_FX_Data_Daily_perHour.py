import pandas as pd
import numpy as np
import xlwings as xw

path_read = "H:\DBAG\FT\900 Work Areas\Robert\\"+"39"+"LiquidtyCME_Eurex_FXpairs_Andreas\CSV_Data_Update"

cme = pd.read_csv(path_read+"\\"+"cme_spreads_Q12022_v6.csv")

eurex = pd.read_csv(path_read+"\\"+"eurex_spreads_Q12022_v6.csv")

eurex.head()
cme.head()
#convert timestamps to datetime objects
eurex.TS_STRING = pd.to_datetime(eurex.TS_STRING)
cme.TS_STRING = pd.to_datetime(cme.TS_STRING)
#drop first row as this is irrelevant
eurex.drop([0], inplace=True)

#Set up the dumping file
book = xw.Book()

for currencypair in eurex.iloc[:,6].unique():
    book.sheets.add(currencypair)

book.sheets["Tabelle1"].delete()
book.save(path_read+"\\"+"SpreadsCurrencyPairs_CME_EurexUpdate.xlsx")
book.close()


#Eurex Data and manual checking
product = "EURGBP"
day = 20220228
eurex[(eurex.iloc[:,6] == product) & (eurex.iloc[:,7] == day)]

#Eurex Data 
for product in eurex.iloc[:,6].unique():
    dMaster = pd.DataFrame(range(0,25), columns = ["Hour"]) #loop trough unique currency pairs
    print(product)
    #for every currency pair loop through every unique day
    for day in eurex.iloc[:,7].unique():
        print(day)
        d = pd.DataFrame(eurex[(eurex.iloc[:,6] == product) & (eurex.iloc[:,7] == day)].iloc[:,14]) #Standard Liquidity Measure
        d.insert(0,"Hour",eurex[(eurex.iloc[:,6] == product) & (eurex.iloc[:,7] == day)].iloc[:,9].dt.hour) #Hours of the day
        d.columns = ["Hour", str(day)]
        dMaster = dMaster.merge(d, how="left")
        #dMaster.drop_duplicates(inplace=True)
        #dMaster.drop_duplicates(inplace=True, subset = [str(day)])
        print(dMaster.head(50))

    
    bookS = xw.Book(path_read +"\\"+"SpreadsCurrencyPairs_CME_EurexUpdate.xlsx")
    bookS.sheets[product].range("A1").options(index=False).value = "Eurex"
    bookS.sheets[product].range("A2").options(index=False).value = dMaster
    bookS.save()
    bookS.close()


#CME Data and manual checking
product = "EURGBP"
day = 20220118
cme[(cme.iloc[:,5] == product) & (cme.iloc[:,6] == day)]

#CME Data
for product in cme.iloc[:,5].unique():
    dMaster = pd.DataFrame(range(0,25), columns = ["Hour"]) #loop trough unique currency pairs
    print(product)
    #for every currency pair loop through every unique day
    for day in cme.iloc[:,6].unique():
        print(day)
        d = pd.DataFrame(cme[(cme.iloc[:,5] == product) & (cme.iloc[:,6] == day)].iloc[:,13]) #Standard Liquidity Measure
        d.insert(0,"Hour",cme[(cme.iloc[:,5] == product) & (cme.iloc[:,6] == day)].iloc[:,8].dt.hour) #Hours of the day
        d.columns = ["Hour", str(day)]
        dMaster = dMaster.merge(d, how="left")
        #dMaster.drop_duplicates(inplace=True)
        dMaster.drop_duplicates(inplace=True, subset = ["Hour"])
        print(dMaster.head(50))

    
    bookS = xw.Book(path_read +"\\"+"SpreadsCurrencyPairs_CME_EurexUpdate.xlsx")
    bookS.sheets[product].range("A28").options(index=False).value = "CME"
    bookS.sheets[product].range("A29").options(index=False).value = dMaster
    bookS.save()
    bookS.close()
