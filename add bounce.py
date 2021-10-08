import pandas as pd
import numpy as np
import openpyxl as pxl


df = pd.read_csv("BOUNCEFILE")

df2 = pd.read_excel("FILENAME", sheet_name="1")

emails, bounce = df['Email Address'].values.tolist(), df['Bounce Type'].values.tolist()

for index in df2.index:
    number = 0
    for email in emails:
        if df2.loc[index, 'Email'] == email:
            df2.loc[index, '1st email stauts'] = bounce[number] 
        number += 1

print(df2)


writer = pd.ExcelWriter('renewal.xlsx', engine = 'openpyxl')
workbook = writer.book

df2.to_excel(writer, sheet_name="1", index=False)
writer.save()
