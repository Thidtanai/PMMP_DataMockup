import xlsxwriter
import pandas as pd

readDF = pd.read_excel(r'writerTest.xlsx')
df = pd.DataFrame({'Money' : [200, 250, 350]})
