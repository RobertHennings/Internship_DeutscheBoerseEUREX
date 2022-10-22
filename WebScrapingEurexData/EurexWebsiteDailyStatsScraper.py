
#Get the desired daily statistics file from the eurex website

def ScrapeDailyStatsEurex(from_date, to_date, url, path_save):
    #Import all needed libarries
    import time
    import threading
    import win32ui, win32gui, win32com, pythoncom, win32con
    from win32com.client import Dispatch
    import glob
    import pandas as pd
    import os 
    import shutil 
    #Set the daily date range for the specified date range
    dates = pd.date_range(start = from_date, end=to_date)
    #Save the years in a designated variable
    year = dates.year.astype("str")
    #Save the month in a designated variable
    month = dates.month.astype("str")
    #Save the days in a designated variable
    day = dates.day.astype("str")
    #Edit the days since we need a zero in front of single number days
    day2 = []
    #put a "0" in front of single number days
    for i in day:
        if len(i) ==1:
            day2.append(i.replace(i,"0"+i))
        else:
            day2.append(i)

    #loop through the time range and retrieve the single files
    for y,m,d in zip(year,month,day2): 

        class IeThread(threading.Thread):
            def run(self):
                pythoncom.CoInitialize()
                ie = Dispatch("InternetExplorer.Application")
                ie.Visible = 0
                ie.Navigate(url+y+m+d+".xls")

        def PushButton(handle, label):
            if win32gui.GetWindowText(handle) == label:
                win32gui.SendMessage(handle, win32con.BM_CLICK, None, None)
                return True

        IeThread().start()
        time.sleep(3)  # wait until IE is started
        wnd = win32ui.GetForegroundWindow()
        if wnd.GetWindowText() == "File Download - Security Warning":
            win32gui.EnumChildWindows(wnd.GetSafeHwnd(), PushButton, "&Save");
            time.sleep(1)
            wnd = win32ui.GetForegroundWindow()
        if wnd.GetWindowText() == "Save As":
            win32gui.EnumChildWindows(wnd.GetSafeHwnd(), PushButton, "&Save");
        print("File downloaded: ", "dailystat_"+y+m+d+".xls")

        #look into the downloads folder and remove the files there into another desired folder
        #finally removing the single files from the Downloads folder to the desired saving path
        shutil.move("Downloads\\dailystat_"+y+m+d+".xls", path_save+"dailystat_"+y+m+d+".xls")
        



ScrapeDailyStatsEurex(from_date = "09-01-2022", to_date = "09-10-2022",url='https://www.eurex.com/resource/blob/3280506/139138bcd179f1cbd5ee8323e7de6fb4/data/dailystat_',path_save = "H:\DBAG\FT\900 Work Areas\Robert\\")



