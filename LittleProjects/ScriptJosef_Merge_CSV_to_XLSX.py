
def main():
    import pandas as pd
    import xlwings as xw
    import glob

    #Change the path to where all the csv data files are located
    path_read= "H:\DBAG\FT\900 Work Areas\Josef\#ETF Options Research\Bloomberg_CSV_Data"

    df_master = pd.DataFrame()
    for i in glob.os.listdir(path_read):
        if "csv" in i: #watch out for the correct ending, here we want the csv files, you could change it to excel for eg.
            df = pd.read_csv(path_read+"\\"+i)
            df.to_csv(path_read+"\\"+i.split(".")[0]+".csv")
            df_master = df_master.append(df)
            print("Success: ",i)
        else:
            pass 
    #save all the data into the same path weher the single data files are located, change file type if needed and pandas write method accordingly       
    df_master.to_excel(path_read+"\\"+"AllData.xlsx")

if __name__ =="__main__":
    main()