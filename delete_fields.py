import pandas as pd
import os

#open the csv file
filename ='S:/LMC/Clinical Research/Submissions/Submissions Database 2.0 BACKEND/_TEMPDATA/research abstract and publication submission form.csv'
f = pd.read_csv(filename)
#select only the columns we want
f = f[[col for col in f.columns if col.startswith("Q") or col in ['V' + n for n in ['1','4','5','8']]]]
##rename them
f = f.rename(columns={'V1':'ResponseID','V4':'Q174','V5':'Q173','V8':'EndDate'})
#drop the first three rows (this is for non-legacy format)
f = f.drop(f.index[[0,1,2]])
f.to_csv(filename, index=False)
