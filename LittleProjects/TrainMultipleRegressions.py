#Read in all the cleaned volume data from each contract, fit the ARIMA(10,2,10) model on it and save some results 
#Import the necessary libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt 
import xlwings as xw
import warnings 
import glob


def main():
    #Save the file path and the filename where the raw monthly data for the single cntracts is located in
    #you might have to edit this for your own use
    path_read = "H:\DBAG\FT\900 Work Areas\Robert\\"
    """
    The file Model_Data should be structured as follows: For each single Product that should be forecasted, include all the regression variables with their normal names as headers
    The variable that should be forecasted should have the same name as the single sheet that the data is included in
    """
    file_name_read = "Model_Data.xlsx"

    file_name_write = "Model_Summary.xlsx"


    
    #Saving the Regression results into a new excel file, each single product will get its own output sheet
    #Save the sheet names to later loop through them and create a regression model for them
    book = xw.Book(path_read+"\\"+file_name_read) 
    #save the order in a list called sheets
    sheets = [i.name for i in book.sheets]
    book.close()
    print(sheets)
    #As the sheet names are now saved we can go on with our regression
    #Open a new book
    new_book = xw.Book()
    #apply the same structure here (e.g. each sheet is one contract that holds the reuslts of the applied analysis)
    for k in sheets:
        new_book.sheets.add(k)
    #for beauty purposes: delete the standard sheet called Tabelle1 thats added, depending on your excel version this mght be called Sheet1, so you might have to edit that in the code here to avoid failures
    new_book.sheets["Tabelle1"].delete()
    #save the file with the empty sheets
    new_book.save(path_read+"\\"+file_name_write)
    #close the file
    new_book.close()

    #Now our Output file is created and holds the same structure of the initial input file that holds the data 
    #The single sheets now need to be filled with the summary of the single regression models
   
    for table in sheets:
        warnings.simplefilter('always', category=UserWarning)
        book = xw.Book(path_read+"\\"+file_name_read)
        #read in the data 
        data = pd.DataFrame(book.sheets[table].range("A1").options(expand="table").value)
        #set the new column names 
        data.columns = data.loc[0]
        #drop the first row
        data.drop([0], inplace=True)
        #replace the NaNs from Bloomberg
        data.replace("#N/A N/A", np.nan, inplace=True)
        #drop all NaNs to have cleaned data
        data.dropna(inplace=True)
        
        y = data[table].copy()
        data.drop(table,axis=1, inplace=True)

        dates = data.iloc[:,0]
        data.drop(data.columns[0], axis=1, inplace=True)
        
        X = data.copy()
        X = X.astype("float64", copy=False)
        y = y.astype("float64", copy=False)
        y = y.values 
        #Now the data is cleaned and ready to run the regression 
        
        from statsmodels.api import OLS
        regression_model = OLS(np.asarray(y),X)
        reg_model_finished = regression_model.fit()
        print("Successss Regression")
        #Save some model results in a dataframe that will be pasted in the output file
        Results = pd.DataFrame(reg_model_finished.fittedvalues,columns=["Fitted Model Values"])
        Results["Original Values"] = y
        
        results_book = xw.Book(path_read+"\\"+file_name_write)
        X.insert(0,"Dates", dates)
        Results.insert(0,"Dates", dates)
        results_book.sheets[table].range("A1").options(index=False).value = X
        results_book.sheets[table].range("L1").options(index=False).value = Results 
        results_book.sheets[table].range("P1").options(index=True).value = reg_model_finished.summary().tables[0].data
        results_book.sheets[table].range("P13").options(index=True).value = reg_model_finished.summary().tables[1].data
        results_book.sheets[table].range("U1").options(index=True).value = reg_model_finished.summary().tables[2].data
        results_book.sheets[table].range("Z1").value = "Model Variable"
        results_book.sheets[table].range("AA1").value = "Coefficients"
        results_book.sheets[table].range("Z2").options(index=True).value = reg_model_finished.params
        results_book.save()
        results_book.close()
        

if __name__ == "__main__":
    main()

table ="OGBM"
sheets 
