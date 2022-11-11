import pandas as pd
import numpy as np 
import matplotlib.pyplot as plt 
import scipy 


#Set the according paths to the needed files
file_path_read = "H:\DBAG\FT\900 Work Areas\Robert\\"+"17"+"SlideUpdatePPTQurterlyReport"
#Provide the file path 
file_name_read = "Gamma_values.xlsx"

#Read in the Gamma values
d = pd.read_excel(file_path_read+"\\"+file_name_read)
#get an overview of the data
d.head()
d.tail()
#Save a copy of the data and work on that version 
d2 = d.copy()
#Drop every column not needed
d2.drop(["NET_POSITION", "UNDERLYING", "VOL", "INTEREST_RATE", "SETTLEMENT", "CALL_PUT_FLAG", "STRIKE", "EXPIRY", "SERIES_ID"], axis=1, inplace=True)
#Check if everything worked out
d2.head()
#Group by Product and Role and save it in the provided path 
d2.groupby(by=["PRODUCT", "ROLE"]).sum().to_excel(file_path_read+"\\"+"GammaAggreagted.xlsx")


d2[d2.PRODUCT=="OGB1"]

dir(np)

d2.iloc[19575,3].dtype

d2.gamma = d2.gamma.astype("str").replace("inf", np.nan).astype("float64")

d2.dropna(inplace=True)
d2.shape


d.FACT_DATE.unique()

d.head()

d["StrikeDistance"] = d.UNDERLYING - d.STRIKE

d.StrikeDistance.plot(kind="kde")
plt.show()



d.StrikeDistance.quantile([0.25,0.50,0.75])

d[(d.PRODUCT == "OGBL") & (d.ROLE == "M")].StrikeDistance.plot(kind="kde")
plt.show()

plt.plot(scipy.stats.norm.pdf(d[(d.PRODUCT == "OGBL") & (d.ROLE == "M")].StrikeDistance))
plt.show()

plt.plot(scipy.stats.norm.cdf(d[(d.PRODUCT == "OGBL") & (d.ROLE == "M")].StrikeDistance))
plt.show()

dir(scipy.stats)

help(scipy.stats.norm)

densit = pd.DataFrame(scipy.stats.norm.pdf(d[(d.PRODUCT == "OGBL") & (d.ROLE == "M")].StrikeDistance),d[(d.PRODUCT == "OGBL") & (d.ROLE == "M")].StrikeDistance)
densit.to_excel(file_path_read+"\\"+"Denisty.xlsx")

d[(d.PRODUCT == "OGBL") & (d.ROLE == "M")].StrikeDistance.quantile([0.25,0.50,0.75])
d[(d.PRODUCT == "OGBL") & (d.ROLE == "M")].StrikeDistance.describe()
densit.head()

d[(d.PRODUCT == "OGBL") & (d.ROLE == "P")].StrikeDistance.describe()
d[(d.PRODUCT == "OGBL") & (d.ROLE == "P")].STRIKE.plot.kde()
plt.vlines(x=d[(d.PRODUCT == "OGBL") & (d.ROLE == "P")].UNDERLYING.mean(), ymin=0, ymax=0.07)
plt.show()



