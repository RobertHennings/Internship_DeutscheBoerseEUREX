#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct 13 21:46:26 2022

@author: Robert_Hennings
"""
#This script serves to scrape daily trading statistics from the EUREX website 
#enter the desired time frame to download the single data files into the desired folder

def main():
    def LoadDailyStatsEurex(from_date, to_date, url,path_save):
        #format fo the dates: "mm-dd-yyyy"
        #import the necessary packages
        import pandas as pd
        import urllib
        
        dates = pd.date_range(from_date,to_date)
        years = dates.year.astype("str").to_list()
        months = dates.month.astype("str").to_list()
        days = dates.day.astype("str")
        #Now edit the months and the dates variable so that if theres a single number, a 0 is put in front of it
        months2 = []
        days2 = []
        
        for m in months:
            if len(m) == 1:
                months2.append("0"+m)
            else:
                months2.append(m)
        for d in months:
            if len(m) == 1:
                days2.append("0"+d)
            else:
                months2.append(d)
        
        for y,m,d in zip(years, months2, days2):
            print("Loading file: ", "DailyStats"+y+m+d)
            urllib.request.urlretrieve(url+y+m+d+".xls", path_save+"DailyStats"+y+m+d+".xls") 
        
        
        
    LoadDailyStatsEurex(from_date = "09-09-2022", 
                        to_date = "09-10-2022", 
                        url="https://www.eurex.com/resource/blob/3280506/139138bcd179f1cbd5ee8323e7de6fb4/data/dailystat_",
                        path_save="/Users/Robert_Hennings//")

if __name__ == "__main__":
    main()