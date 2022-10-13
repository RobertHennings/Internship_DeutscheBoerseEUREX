#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct 13 21:46:26 2022

@author: Robert_Hennings
"""

import urllib
#Download Daily Trading Statistics from EUREX Website
path_save = "/Users/Robert_Hennings/Downloads//"
website = "https://www.eurex.com/resource/blob/3280506/139138bcd179f1cbd5ee8323e7de6fb4/data/dailystat_"

dates = ["20221012", "20221011"]

for date in dates:
    url = website+date+".xls"
    urllib.request.urlretrieve(url, path_save+url[-12:])  
    
    




