import xlsxwriter
import pandas as pd

df = pd.DataFrame({'Money' : [200, 250, 350]})
writer = pd.ExcelWriter('writerTest.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name="Page 1", index=False)
writer.save()
