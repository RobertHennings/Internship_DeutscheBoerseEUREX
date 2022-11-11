def main():
    import pandas as pd 
    import xlwings as xw
    import numpy as np
    path = "H:\DBAG\FT\900 Work Areas\Robert\\"+"16"+"PythonScripts"
    file_name = "ZinsesZinsOutputNeu.xlsx"
    path_save = path+ "\\"+file_name
    
    def Zinseszins(startkapital, ZinsJahr, jahre):
        alle = []
        for i in ZinsJahr:
           print(f"{np.round(i*100,3)}"+" % Zins")
           kapital = []
           for j in range(0,jahre+1):
                ertrag = startkapital*(1+i)**(j)
                kapital.append(ertrag)
        
           globals()["Kapital"+f"{np.round(i*100,3)}"+"%"] = pd.DataFrame(kapital)
           alle.append("Kapital"+f"{np.round(i*100,3)}"+"%")
        globals()["alle"] = alle
        
    Zinseszins(1, [0.07, 0.05, 0.02, 0.01, 0.005], 2000)
    
    def WriteDataTable(table_name, where):
        book_new.sheets.add(table_name) #Set the Sheet name to exactly the original ones from the main file
        book_new.sheets[table_name].range(where).value = globals()[f"{table_name}"]
        
    #excel_app =xw.App(visible=False)
    #book_new= excel_app.books.add()
    
    book_new = xw.Book()
    #Setting the new opened file as the active one in utilization
    book_new = xw.books.active
    
    Gesamt = pd.DataFrame()
    for i in alle:
       Gesamt =  pd.concat([Gesamt, globals()[f"{i}"]], axis=1)
    
    Gesamt.columns = alle 
    Gesamt.index.name = "Jahre"
    globals()["Gesamt"] = Gesamt
    alle.append("Gesamt")
    for i in alle:
        WriteDataTable(i, "A1")
        
    book_new.sheets[book_new.sheets.count-1].delete()
    book_new.save(path_save)

if __name__ == "__main__":
    main()