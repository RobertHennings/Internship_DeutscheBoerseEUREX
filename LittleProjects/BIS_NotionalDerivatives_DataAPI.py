# -*- coding: utf-8 -*-
"""
Created on Thu Aug  4 15:57:01 2022

@author: ka710
"""

#####################################################################################################################################################################
#Reading in the single downloaded data files from the BIS website
import glob
path_read = "H:\DBAG\FT\900 Work Areas\Robert\\"+"10"+"InterestRateDerivativesNotionalamountsOutstanding"   

glob.os.listdir(path_read)

#Reading in all the single files, write them into one excelfile (one sheet per file) and one master sheet that holds all the data tables

excel_app =xw.App(visible=False)
book_new= excel_app.books.add()

l =[]
for i in glob.os.listdir(path_read):
    l.append(i.split(".")[0])
l.sort()
# l.sort(reverse=True)

for i in l:
    if not "table" in i:
        l.pop(l.index(i))
#use this version of the loop if the single data tables are read in with the glob.os path
# for i in glob.os.listdir(path_read):
#     if ".xlsx" in i:
#         df = pd.read_excel(path_read+"\\"+i, header = 7)
#         book_new.sheets.add(i.split(".")[0])
#         book_new.sheets[book_new.sheets == i.split(".")[0]].range("A1").value = df
#         i == glob.os.listdir(path_read)[len(glob.os.listdir(path_read))-1]
#         print("Finished")

#     else:
#         pass
    
#different version with sorted list of single data files
for i in l:
    df = pd.read_excel(path_read+"\\"+i+".xlsx", header = 7)
    book_new.sheets.add(i)
    book_new.sheets[book_new.sheets == i].range("A1").value = df
    print("Finished")

#deleting default sheet that is generated when opening a new excel file called Sheet1
book_new.sheets["Sheet1"].delete()

book_new.save(path_read+"\\"+"AllDataFiles.xlsx") #saving the Export File at the desired file path with the desired name
book_new.close() #closing the file 

#Opening the saved file and extracting the needed data
#Three Graphics for USD: IRS, CDS, Total Options and three Graphics for EUR: IRS, CDS, Total Options
#Reading in the File that houses all single data tables on a HY basis

book = xw.Book(path_read+"\\"+"AllDataFiles.xlsx")

#Extracting all the sheet names from the file
s = []
for i in book.sheets:
    s.append(i.name)

#USD
#FRAs: Cell D5 in every sheet
#SWAPS: Cell D10 in every sheet


#EUR
#FRAs: Cell E5 in every sheet
#SWAPS: Cell E10 in every sheet

#Loop through every sheet and extract the data
#Create an empty dataframe thats gonna be filled with the values
USD = pd.DataFrame(index = range(book.sheets.count), columns = ["HY","FRA", "Swaps", "Options"])
EUR = pd.DataFrame(index = range(book.sheets.count), columns = ["HY","FRA", "Swaps", "Options"])

ind = 0
for j in book.sheets:
    USD.iloc[ind,0] = j.range("B1").value
    USD.iloc[ind,1] = j.range("D5").value
    USD.iloc[ind,2] = j.range("D10").value
    USD.iloc[ind,3] = j.range("D15").value
    EUR.iloc[ind,0] = j.range("B1").value
    EUR.iloc[ind,1] = j.range("E5").value
    EUR.iloc[ind,2] = j.range("E10").value
    EUR.iloc[ind,3] = j.range("E15").value
    ind +=1


book.sheets.add("FRA_Swap_Options_USD", before=book.sheets[0])
book.sheets.add("FRA_Swap_Options_EUR", after=book.sheets[0])

book.sheets[0].range("A1").value = USD
book.sheets[1].range("A1").value = EUR

book.save(path_read+"\\"+"AllDataFiles.xlsx")
book.close()
