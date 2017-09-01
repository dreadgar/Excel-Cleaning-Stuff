# -*- coding: utf-8 -*-
"""To add and delete columns in a CSV file"""
import pandas as pd
import numpy as np
from datetime import date, timedelta 



df = pd.read_csv(r'C:\Users\accou\Desktop\Aging Beta 10.11.16.csv') #calling csv file

df['Sent On'] = df['Sent On'].str.replace('|','') 
#removes formula bars-- can also be used to remove any character/s in a string
df['Sent On'] = pd.to_datetime(df['Sent On']) 
#Converts date and time "object" to date/time format

df['Last Activity'] = df['Last Activity'].str.replace('|','') #removes formula bars
df['Last Activity'] = pd.to_datetime(df['Last Activity']) #Converts to date/time format

#setting target date as 4 days ago
target_date = pd.to_datetime('today') - timedelta(4)
#removing all entries less than 4 days old
df = df[(df['Sent On'] <= target_date)]


#deleting unnecessary columns
del df['Remaining Signatures']
del df['Status']
del df['Sender User ID']

#Trimming repetative non-uniform data from in front of desired data
df['Signer List'] = df['Signer List'].apply(lambda x: x.split('Customer - ')[1])
#adding column
df['Rep'] = pd.Series(index=df.index)
#Pulling Rep initials from subject line using ") " as the marker
df['Rep'] = df['Subject'].apply(lambda x: x.split(') ')[1])
#Sorts data by Rep alphabetically
df = df.sort_values('Rep')
#Renaming columns
df = df.rename(columns={'Subject': 'Document', 'Signer List': 'Contact'})

#print(df.dtypes) # (commented out) used to check data types




#Creating seperate dataframes for individual reps
dfbb = df[(df['Rep'] == 'BB')]
dfdp = df[(df['Rep'] == 'DP')]
dfkw = df[(df['Rep'] == 'KW')]


#creating a date (today) to use as a version date with specific formatting
version_date = []
today = date.today()
version_date.append(today)
#double-check version date and formatting (commented out)
#print(version_date[0])


# Create a Pandas Excel writer using XlsxWriter as the engine and the version date from the previous statements.
writer = pd.ExcelWriter((r'C:\Users\accou\Desktop\Docusign Aging Reports\Aging_beta_' + str(version_date[0]) + '.xlsx'), engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save() 

#Repeating for each rep subreport
writer = pd.ExcelWriter((r'C:\Users\accou\Desktop\Docusign Aging Reports\Aging_beta_for_BB_' + str(version_date[0]) + '.xlsx'), engine='xlsxwriter')
dfbb.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()

writer = pd.ExcelWriter((r'C:\Users\accou\Desktop\Docusign Aging Reports\Aging_beta_for_DP_' + str(version_date[0]) + '.xlsx'), engine='xlsxwriter')
dfdp.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()

writer = pd.ExcelWriter((r'C:\Users\accou\Desktop\Docusign Aging Reports\Aging_beta_for_KW_' + str(version_date[0]) + '.xlsx'), engine='xlsxwriter')
dfkw.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()
