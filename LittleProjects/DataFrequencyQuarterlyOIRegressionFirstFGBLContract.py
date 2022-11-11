# -*- coding: utf-8 -*-
"""
Created on Thu Aug  4 16:17:39 2022

@author: ka710
"""
import pandas as pd
import glob
import seaborn as sns
import matplotlib.pyplot as plt
import scipy
import numpy as np
from scipy import stats
import math
from matplotlib import pyplot
from statsmodels.graphics.gofplots import qqplot
#Documentation of the whole Data Processing and all steps for the Open Interest Regression to forecast revenues 
#1) Dowload the Open Interest Data for the FGBL and all its open contracts from the EUREX Database
#2) Split the Dataset into one single for each contract (1,2,3)
#3) Take the first contract data
#4) Take the Duration Data from Bloomberg for each Contract that was the actual open one for each day
#5) Compute the Duration Adjusted Open Interest and the Pct Change of the Open Interest on a daily frequency
#6) Compute the Indicator column which states when the actual CTD contract is switched over to the next one
#7) Resample the daily sequence to a quarterly one 
#8) Compute the Ratio of Dur Qt/Dur Qt-1
#9) Do the substraction: Substract the Dur Qt/Dur Qt-1 from the column "PercentageChangeOIFirst" to get the EQ (Error term)

#Import the basic data on a daily frequency containing the Percentage Change of OIof the first Contract, the Duration Adjusted OI and the Switching Indicator
path_read = "H:\DBAG\FT\900 Work Areas\Robert\\"+"04"+"RegressionVolumesRevenues\RegressionOpenInterest\Reg_V10_OI_RH"   
file_name_read = "Variables.xlsx"

path_save = "H:\DBAG\FT\900 Work Areas\Robert\\"+"04"+"RegressionVolumesRevenues\RegressionOpenInterest\Data"
file_name_write = "QuarterlyFGBL_1Contract"

#The Duration from the first contract
DurationFirst = pd.read_excel("H:\DBAG\FT\900 Work Areas\Robert\\"+"04"+"RegressionVolumesRevenues\RegressionOpenInterest\Data\DataRegressionOpenInterestVersions.xlsx", 
                              sheet_name ="DurationFuture", 
                              header = 21,index_col=0, 
                              usecols=["Date","RX1 Comdty"])


#The Open Interest from the first contract
OpenIntFirst = pd.read_excel("H:\DBAG\FT\900 Work Areas\Robert\\"+"04"+"RegressionVolumesRevenues\RegressionOpenInterest\Data\FGBL_ExpNo1_OI.xlsx", 
                             index_col=0, 
                             usecols=["Date","OI"], parse_dates=True)


OGdataDailyFirstContract = pd.read_excel(path_read+"\\"+file_name_read, index_col=0)
OGdataDailyFirstContract.head()

#Resample the daily frequency to a quarterly one, taking the averages of the daily values
dataQuarterlyFisrtContract = OGdataDailyFirstContract.resample("Q").mean()

#Just for dumping data into the file system
# from pandas import ExcelWriter
# with ExcelWriter(path_save+"\\"+file_name_write+".xlsx") as writer:
#   dataQuarterlyFisrtContract.to_excel(writer, sheet_name="QuarterlyData")
#   OGdataDailyFirstContract[OGdataDailyFirstContract.Indicator_RX1 == 1].to_excel(writer, sheet_name="FirstContractSwitchingDay")


#Substract the Dur Qt/Dur Qt-1 from the column "PercentageChangeOIFirst" to get the EQ (Error term)
#Compute Dur Qt/Dur Qt-1
DiffDur = (dataQuarterlyFisrtContract.iloc[:,1]/dataQuarterlyFisrtContract.iloc[:,1].shift(1))
DiffDur.dropna(inplace=True)

#Now do the substraction

EQ = dataQuarterlyFisrtContract.iloc[1:,0] - DiffDur

EQ
#Now the error term is computed and analysis can be carried out
#Inspect the distribution of the term, its expected value, its std, its kurtosis
#kurtosis should be between 1 and 25
#inspect tests on normality like kolmogorov-smirbov and shapiro wilks
#compute onesided p tests, two sided p test and t statistics
EQ.mean()
EQ.std()
EQ.kurt()
EQ.describe()

plt.hist(EQ)
plt.show()

qqplot(EQ, line="s")
sns.distplot(EQ)

def T_test(mu_stich, mu_hyp, std_stich, n_stich):
    """
    Parameters
    ----------
    mu_stich : TYPE
        mean of the sample 
    mu_hyp : TYPE
        theoretical mean 
    std_stich : TYPE
        standard deviation of the sample
    n_stich : TYPE
        sample size

    Returns
    -------
    t_stat : TYPE
        t statistic of the sample

    """
    t_stat = (mu_stich-mu_hyp)/((std_stich)/math.sqrt(n_stich))
    
    print("Value of the t-statistic: ", t_stat)
    return t_stat 

T_test(EQ.mean(),0,EQ.std(), len(EQ))
    
from scipy import stats
def one_sample_one_tailed(sample_data, popmean, alpha=0.05, alternative='greater'):
    t, p = stats.ttest_1samp(sample_data, popmean)
    print ('t:',t)
    print ('p:',p)
    if alternative == 'greater' and (p/2 < alpha) and t > 0:
        print ('Reject Null Hypothesis for greater-than test')
    if alternative == 'less' and (p/2 < alpha) and t < 0:
        print ('Reject Null Hypothesis for less-thane test')

#H0: Sample comes from a normal distribution
#H1: Sample comes from an other Distribution
#If p < 0.05 H0 has to be rejected -> not normal distributed

#General t test
sample_data = EQ      
one_sample_one_tailed(sample_data,0)  

#left sided test
stats.t.sf(abs(-8.931068045645063), df=len(EQ)-1)
#two sided test
stats.t.sf(abs(-8.931068045645063), df=len(EQ)-1)*2


#Look up the critical values for the t statistic
#two tailed critical value, so take alpha/2 -> 1-alpha/2
stats.t.ppf(q=0.975, df=len(EQ)-1)

#one tailed critical value
stats.t.ppf(q=0.05, df=len(EQ)-1)

#two tailed p value
2*(1-stats.t.cdf(x=-8.931068045645063, df=len(EQ)-1))

#one tailed p value
1-stats.t.cdf(x=-8.931068045645063, df=len(EQ)-1)

#Perfrom tests on the distribution of the Error term, see if it is normal distributed
#Shapiro Wilks test
stats.shapiro(EQ)

stats.shapiro(EQ).pvalue<0.025

#Kolmogorov Smirnov Test
stats.kstest(EQ,stats.norm.cdf)

stats.kstest(EQ,stats.norm.cdf).pvalue < 0.025

#Simple normal test
stats.normaltest(EQ)
stats.normaltest(EQ).pvalue < 0.025



dir(stats.norm)
help(stats.norm.cdf)
stats.norm.cdf(-8.03)




######################################################################################################################################################################
import datetime as dt
#Set up a new Version, taking the Dur Q2 / Dur Q1 as simple ratio instead of the Duration adjusted Open Interest
#using monthly data frequency directly from Bloomberg

path_read = "H:\DBAG\FT\900 Work Areas\Robert\\"+"04"+"RegressionVolumesRevenues\FirstContract"   
file_name_read = "OI_Volume_Dur_MonthlyFirstContract.xlsx"
DataMonthlyFirst = pd.read_excel(path_read+"\\"+file_name_read, header=21, index_col=0)

#Computing the needed ratios from the monthly frequency
DataMonthlyFirst["OI_Change"] = DataMonthlyFirst.iloc[:,1].pct_change()

DataMonthlyFirst["DUR_Change"] = DataMonthlyFirst.iloc[:,0] / DataMonthlyFirst.iloc[:,0].shift(1)

DataMonthlyFirst["EQ_Diff"] = DataMonthlyFirst["OI_Change"] - DataMonthlyFirst["DUR_Change"]

#Drop all NAs
DataMonthlyFirst.dropna(inplace=True)

#Have a first look at the data
plt.hist(DataMonthlyFirst["EQ_Diff"])
plt.show()

qqplot(DataMonthlyFirst["EQ_Diff"], line="s")

sns.distplot(DataMonthlyFirst["EQ_Diff"])

#Perfrom some normality tests
sample_data = DataMonthlyFirst["EQ_Diff"]
one_sample_one_tailed(sample_data,0) 



stats.normaltest(DataMonthlyFirst["EQ_Diff"])


stats.kstest(DataMonthlyFirst["EQ_Diff"],stats.norm.cdf)



#Version that Leon proposed: take monthly data frequency, take a look at the middle of the previous month before a roll month and take the OI up to to that point and substract it from the from the volume in the roll month

#take daily data, look at the previous months of roll months, sum up the OI up until the middle of this month, subtract this value from the trading volume in the roll month

roll_months = ["March", "June", "September", "December"]
#sum up the OI until the middle of the months of February, May, August, November from the daily data
#on a monthly scale substract the monthly summed up OI from the monthly trading volume
OpenIntFirst.index = pd.to_datetime(OpenIntFirst.index)

OpenIntFirst.index[0].month_name()

indices_to_sum = []
def ComputeSumRollMonth(series_to_sum,DatetimeInd,roll_months,when):
    months_when_to_sum = ["February", "May", "August", "November"]
    
    for i,j in zip(series_to_sum,DatetimeInd):
        if j.month_name() in months_when_to_sum and j.day <when:
            indices_to_sum.append(j)
    
ComputeSumRollMonth(OpenIntFirst.iloc[:,0],OpenIntFirst.index, roll_months,15)
    
df = pd.DataFrame()
for f in indices_to_sum:
    # print(OpenIntFirst.loc[f,:])
    df = df.append(OpenIntFirst.loc[f,:])

OI_beforeRollMonth = df.resample("M").sum()

DataMonthlyFirst.tail()
df2 = pd.DataFrame([0], columns = ["OI"])
df2["t"] = pd.to_datetime("2022-06-30")

df2.set_index(df2["t"], inplace=True)
df2.index_name = "Index"
df2.drop(["t"], axis=1, inplace= True)


OI_beforeRollMonth = OI_beforeRollMonth.append(df2)


s1  = OI_beforeRollMonth.iloc[:,0].values
DataMonthlyFirst["SuntractedTradingVolume"] = DataMonthlyFirst.PX_VOLUME.values - OI_beforeRollMonth.OI.values
DataMonthlyFirst["OIbeforeRoll"] = OI_beforeRollMonth.OI
DataMonthlyFirst.index[0]

#The colum SuntractedTradingVolume holds the trading volume minus the sum of the OI of the previous month
DataMonthlyFirst.to_excel(path_read+"\\"+"EditedVolumeDataNoRollEffect.xlsx")



######################################################################################################################################################################
#Use the paper: Economic determinants of trading volume in future markets
#First test if thh total trading volume on a monthly basis follows a random process

dir()
