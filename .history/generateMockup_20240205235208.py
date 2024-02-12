import xlsxriter 
import pandas as pd

fileName = 'Student_data_2_2566.xlsx'
medDF = pd.read_excel(fileName, 13)

print(medDF)