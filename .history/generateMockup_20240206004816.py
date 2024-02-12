import xlsxwriter
import pandas as pd

fileName = 'Student_data_2_2566.xlsx'
medReadDF = pd.read_excel(fileName, 13)
dentReadDF = pd.read_excel(fileName, 14)
medTechReadDf = pd.read_excel(fileName, 15)
NurseReadDf = pd.read_excel(fileName, 16)
frames = [medReadDF]
result = pd.concat(frames)

writer = pd.ExcelWriter('writerTest.xlsx', engine='xlsxwriter')

result.to_excel(writer, sheet_name="sheet 1", index=False)
writer.save()
