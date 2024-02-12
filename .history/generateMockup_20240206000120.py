import xlsxwriter
import pandas as pd

fileName = 'Student_data_2_2566.xlsx'
medReadDF = pd.read_excel(fileName, 13)
medFrames = [medReadDF]
result = pd.concat(medFrames)

writer = pd.ExcelWriter('writerTest.xlsx', engine='xlsxwriter')

result.to_excel(writer, sheet_name="sheet 1", index=False)
writer.save()
