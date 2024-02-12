import xlsxwriter
import pandas as pd
import random as rnd

fileName = 'Student_data_2_2566.xlsx'
medReadDF = pd.read_excel(fileName, 13)
dentReadDF = pd.read_excel(fileName, 14)
medTechReadDf = pd.read_excel(fileName, 15)
NurseReadDf = pd.read_excel(fileName, 16)
frames = [medReadDF, dentReadDF, medTechReadDf, NurseReadDf]
result = pd.concat(frames)
rnd.shuffle(result)

writer = pd.ExcelWriter('writerTest.xlsx', engine='xlsxwriter')

result.to_excel(writer, sheet_name="sheet 1", index=False)
writer.save()
