import pandas as pd
import os

#open the csv file
dir = os.path.dirname(__file__)
filename ='S:/LMC/Clinical Research/Submissions/Submissions Database 2.0 BACKEND/_TEMPDATA/research abstract and publication submission form'
f = pd.read_csv(filename + ".csv")
#select only the columns we want
f = f[[col for col in f.columns if col.startswith("Q") or col in ['V' + n for n in ['1','4','5','8']] or col == 'ResponseId']]
##rename them
f = f.rename(columns={'V1':'ResponseID','V4':'Q174','V5':'Q173','V8':'EndDate'})
#drop the first row
f = f.drop(f.index[[0,1,2]])

#convert phone numbers and pagers to string
f['Q311'] = f['Q311'].apply(str)
f['Q312'] = f['Q312'].apply(str) 
f['Q321'] = f['Q321'].apply(str) 
f['Q322'] = f['Q322'].apply(str)

#delete 'nan's from phone number fields
f.loc[f['Q311'] == 'nan','Q311'] = '--'
f.loc[f['Q312'] == 'nan','Q312'] = '--'
f.loc[f['Q321'] == 'nan','Q321'] = '--'
f.loc[f['Q322'] == 'nan','Q322'] = '--'

###replace download urls with the correct url
##f['Q53'] = f['Q53'].str.replace('s.qualtrics', 'nyumc.qualtrics')
##f['Q54'] = f['Q54'].str.replace('s.qualtrics', 'nyumc.qualtrics')

#export to csv
f.to_csv(filename + '.csv', index=False)



  
