### This file imports excel data to panda, processes it using several methods,
### and then saves as a different excel file
"""Cleaning Excel Data"""
import pandas as pd
import numpy as np


df = pd.read_excel("Your_excel_sheet.xls") #calling xls file

df['Column'] = df['Column'].str.replace('|','') 
#removes formula bars-- can also be used to remove any character/s in a string
df['Column'] = pd.to_datetime(df['Column']) 
#Converts date and time "object" to date/time format

df['Column'] = df['Column'].str.replace('|','') ##\removes formula bars
df['Column'] = pd.to_datetime(df['Column']) #Converts to date/time format

#deleting unnecessary columns
del df['Column1']
del df['Column2']
del df['Column3']

#Trimming repetative non-uniform data front in front of desired data using a consistent marker common to every line
df['Column4'] = df['Column4'].apply(lambda x: x.split('marker before which everything is deleted')[1])

print(df.dtypes) #used to check data types

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('output_file_name.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.save()


