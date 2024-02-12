import xlsxwriter
import pandas as pd
import random as rnd

# Read file (const)
fileName = 'Student_data_2_2566.xlsx'
# Read selected sheet
medReadDF = pd.read_excel(fileName, 13)
dentReadDF = pd.read_excel(fileName, 14)
medTechReadDf = pd.read_excel(fileName, 15)
NurseReadDf = pd.read_excel(fileName, 16)
# Concat sheet
frames = [medReadDF, dentReadDF, medTechReadDf, NurseReadDf]
connected_frames = pd.concat(frames)
# Shuffle list 
shuffled_result = connected_frames.sample(frac=1).reset_index(drop=True)

# Split data
group_size = len(shuffled_result) // 3
group_1 = shuffled_result.iloc[:group_size]
group_2 = shuffled_result.iloc[group_size:2*group_size]
group_3 = shuffled_result.iloc[2*group_size:]

# Write file (const)
writer = pd.ExcelWriter('writerTest.xlsx', engine='xlsxwriter')
group_1.to_excel(writer, sheet_name="sheet 1", index=False)
group_2.to_excel(writer, sheet_name="sheet 2", index=False)
group_3.to_excel(writer, sheet_name="sheet 3", index=False)
writer.save()

print("Write end")