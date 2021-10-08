import pandas as pd
import numpy as np
import openpyxl as pxl



df2 = pd.read_excel("FILENAME", sheet_name="2")
dfSorted2 = pd.read_excel("FILENAME", sheet_name="2lic")

    
for index in df2.index: 
    df2['Date'] = df2['Expired1'].dt.date
    df2['Date2'] = df2['Expired2'].dt.date
    newRow1 = {
        'Expired' : df2.loc[index, 'Date'],
        'Company' : df2.loc[index, 'Company'],
        'First Name' : df2.loc[index, 'First Name'],
        'Email' : df2.loc[index, 'Email'],
        'Phone' : df2.loc[index, 'Phone'],
        'Date' : df2.loc[index, 'Date'],
        'Product' : df2.loc[index, 'Product1'],
        'QTY' : df2.loc[index, 'QTY1'],
        'License' :df2.loc[index, 'License1']
    }

    newRow2 = {
        'Expired' : df2.loc[index, 'Date2'],
        'Company' : df2.loc[index, 'Company'],
        'First Name' : df2.loc[index, 'First Name'],
        'Email' : df2.loc[index, 'Email'],
        'Phone' : df2.loc[index, 'Phone'],
        'Date' : df2.loc[index, 'Date2'],
        'Product' : df2.loc[index, 'Product2'],
        'QTY' : df2.loc[index, 'QTY2'],
        'License' :df2.loc[index, 'License2']
    }
    data = pd.Series(newRow1, name ='x')
    data2 = pd.Series(newRow2, name ='x')
    dfSorted2 = dfSorted2.append(data, ignore_index=False)
    dfSorted2 = dfSorted2.append(data2, ignore_index=False)

print(dfSorted2)




writer = pd.ExcelWriter('renewal.xlsx', engine = 'openpyxl')
workbook = writer.book

dfSorted2.to_excel(writer, sheet_name="1", index=False)
writer.save()
