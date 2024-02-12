import xlsxwriter
import pandas as pd

readDF = pd.read_excel(r'writerTest.xlsx')
df = pd.DataFrame({'Money' : [200, 250, 350]})
frames = [readDF, df]
result = pd.concat(frames) 

writer = pd.ExcelWriter('writerTest.xlsx', engine='xlsxwriter')

result.to_excel(writer, sheet_name="sheet 1", index=False)
writer.save()