import requests
import zipfile
import pyodbc as db
try: import simplejson as json
except ImportError: import json

lastResponse = True

#First retrieve the last responseID from the Access DB

if lastResponse:
    datafile = 's:/lmc/clinical research/submissions/submissions database 2.0 backend/Submissions Database 2.0_be.accdb'
    connectionString = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};dbq=%s' % datafile
    conn = db.connect(connectionString)
    cursor = conn.cursor()

    cursor.execute('SELECT id FROM (SELECT Max(timestamp) AS MaxTime, Last(responseid) AS id FROM tblsubmissions);')

    for row in cursor.execute('SELECT id FROM (SELECT Max(timestamp) AS MaxTime, Last(responseid) AS id FROM tblsubmissions);'):
        lastResponseId = row.id

    conn.close()
    del conn
    
#Setting user Parameters
apiToken = "YvdlJ7FdzbFdqAZHQxQQMH3oXf4mC3mJAtYdEJm6"
surveyId = 'SV_ahQbp5LCnOJDVHf'
fileFormat = 'csv2013'
useLabels= 'true'

#Setting static parameters
requestCheckProgress = 0
baseUrl = 'https://nyumc.qualtrics.com/API/v3/responseexports/'
headers = {
    'content-type': "application/json",
    'x-api-token': apiToken,
    }

#Creating Data Export
downloadRequestUrl = baseUrl
if lastResponse:
    downloadRequestPayload = '{"format":"' + fileFormat + '","surveyId":"' + surveyId + '","useLabels":' +  useLabels +  ',"lastResponseId":"' + lastResponseId + '"}'
else:
    downloadRequestPayload = '{"format":"' + fileFormat + '","surveyId":"' + surveyId + '","useLabels":' +  useLabels +  '}'
        
downloadRequestResponse = requests.request("POST", downloadRequestUrl, data=downloadRequestPayload, headers=headers)
progressId = downloadRequestResponse.json()['result']['id']
print(downloadRequestResponse.text)




#Checking on Data Export Progress and waiting until export is ready
while requestCheckProgress < 100:
  requestCheckUrl = baseUrl + progressId
  requestCheckResponse = requests.request("GET", requestCheckUrl, headers=headers)
  requestCheckProgress = requestCheckResponse.json()['result']['percentComplete']
  print( "Download is " + str(requestCheckProgress) + " complete")


#Downloading and unzipping file
requestDownloadUrl = baseUrl + progressId + '/file'
r = requests.request("GET", requestDownloadUrl, headers=headers, stream=True)
r.encoding = 'utf16'

with open('Request.zip', 'wb') as f:
    for chunk in r.iter_content(chunk_size=1024):
      f.write(chunk)

zipfile.ZipFile('Request.zip').extractall("TEMPDATA")
zipfile.ZipFile('Request.zip').close


import pandas as pd
import os

#open the csv file
dir = os.path.dirname(__file__)
filename ='S:/LMC/Clinical Research/Submissions/Submissions Database 2.0 BACKEND/TEMPDATA/research abstract and publication submission form'
f = pd.read_csv(filename + ".csv")
###select only the columns we want
##f = f[[col for col in f.columns if col.startswith("Q") or col in ['V' + n for n in ['1','4','5','8']]]]
####rename them
##f = f.rename(columns={'V1':'ResponseID','V4':'Q174','V5':'Q173','V8':'EndDate'})
###drop the first row
##f = f.drop(f.index[[0]])
##
###convert phone numbers and pagers to string
##f['Q311'] = f['Q311'].apply(str)
##f['Q312'] = f['Q312'].apply(str)
##f['Q321'] = f['Q321'].apply(str)
##f['Q322'] = f['Q322'].apply(str)
##
###delete 'nan's from phone number fields
##f.loc[f['Q311'] == 'nan','Q311'] = '--'
##f.loc[f['Q312'] == 'nan','Q312'] = '--'
##f.loc[f['Q321'] == 'nan','Q321'] = '--'
##f.loc[f['Q322'] == 'nan','Q322'] = '--'
##
#####replace download urls with the correct url
####f['Q53'] = f['Q53'].str.replace('s.qualtrics', 'nyumc.qualtrics')
####f['Q54'] = f['Q54'].str.replace('s.qualtrics', 'nyumc.qualtrics')
##
###export to csv
##f.to_csv(filename + '.csv', index=False)
##
##

# Run Access Macro

import win32api,time
from win32com.client import Dispatch

strDbName = 's:/lmc/clinical research/submissions/submissions database 2.0 backend/Submissions Database 2.0_Pre-Deployment.accdb'
objAccess = Dispatch("Access.Application")
#objAccess.Visible = True
objAccess.OpenCurrentDatabase(strDbName)
objDB = objAccess.CurrentDb()
objAccess.runMacro('sync')
#objAccess.Application.Quit()
  
