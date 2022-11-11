# -*- coding: utf-8 -*-
"""
Created on Fri Aug 26 10:28:39 2022

@author: ka710
"""

import pandas as pd
path_read = "H:\DBAG\FT\900 Work Areas\Robert\\"+"17"+"SlideUpdatePPTQurterlyReport\Gamma_values.xlsx"


#Read in file and push for each product one sheet
#for each Product further differentiate between the three Account types M A P

df = pd.read_excel(path_read)

df.PRODUCT.unique()

#Adding a Long short Flag
df["Long_Short"] = "Na"

for i in range(df.shape[0]):
    if df.iloc[i,4] <0:
        df.iloc[i,13] = "Short"
    else:
        df.iloc[i,13] = "Long"

#Set up a dataframe for storing the net gammas per product per Account type
df_NetGamma = pd.DataFrame(index = range(3), columns = df.PRODUCT.unique())


for product in df.PRODUCT.unique():
    globals()[f"{product}_df"] = df[df.PRODUCT == product]
    # del globals()[f"{product}_df"]
    # globals()[f"{product}_df"].reset_index(drop=True, inplace=True)
    
   

#further differentiate after Account Type and sum up all Long Gamma and substract the Short Gamma of it

OOAT_df.loc[(OOAT_df.ROLE == "A")  & (OOAT_df.Long_Short=="Long")].gamma.sum()
OOAT_df.loc[(OOAT_df.ROLE == "A")  & (OOAT_df.Long_Short=="Short")].gamma.sum()




dir(df.PRODUCT.unique())
dir(df.PRODUCT.unique().tolist().insert(0,"Account"))
help(df.PRODUCT.unique().tolist().insert)
df.PRODUCT.unique().tolist().insert(0,"Account")
