import pandas as pd
import os

#open the csv file
dir = os.path.dirname(__file__)
filename ='S:/LMC/Clinical Research/Submissions/Submissions Database 2.0 BACKEND/TEMPDATA/research abstract and publication submission form.csv'
f = pd.read_csv(filename)
#select only the columns we want
f = f[[col for col in f.columns if col.startswith("Q") or col in ['V' + n for n in ['1','4','5','8']]]]
##rename them
f = f.rename(columns={'V1':'ResponseID','V4':'Q174','V5':'Q173','V8':'EndDate'})
#drop the first row
f = f.drop(f.index[[0]])
f.to_csv(filename, index=False)
