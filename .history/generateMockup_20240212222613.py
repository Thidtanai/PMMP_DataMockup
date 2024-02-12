import xlsxwriter
import pandas as pd
import random as rnd
import numpy as np

seed_value = 69
np.random.seed(seed_value)

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
    # group_1 -> Education and Professional Development
    # group_2 -> Leisure and Arts
    # group_3 -> Health, Social, and Personal Growth
group_size = len(shuffled_result) // 3
education_group = shuffled_result.iloc[:group_size]
art_group = shuffled_result.iloc[group_size:2*group_size]
social_group = shuffled_result.iloc[2*group_size:]
# Add event tag
tag_list = [
    "วิชาการ", "วิชาชีพ", "กิจการนักศึกษา", "เพิ่มทักษะ", "เทคโนโลยี",
    "บันเทิง", "ศิลปะ", "การแสดง", "ศาสนา","เทศกาล", 
    "สุขภาพ", "เข้าสังคม", "จิตอาสา", "ท่องเที่ยว", "กีฬา"
]
education_prob = [0.12, 0.12, 0.12, 0.12, 0.12,
                   0.02, 0.06, 0.02, 0.06, 0.06,
                   0.06, 0.06, 0.02, 0.02, 0.02
                   ]
art_prob = [0.06, 0.06, 0.06, 0.06, 0.02,
                   0.12, 0.12, 0.12, 0.12, 0.12,
                   0.02, 0.06, 0.02, 0.02, 0.02
                   ]
social_prob = [0.06, 0.02, 0.06, 0.06, 0.02,
                   0.06, 0.02, 0.02, 0.02, 0.06,
                   0.12, 0.12, 0.12, 0.12, 0.12
                   ]

def addTag(group_size, list, prob):
    tag_col = []
    for i in range(group_size):
        rnd_tag = np.random.choice(list, 3, p=prob)
        tag_col.append(rnd_tag)
        
    print(tag_col[:5])
    

addTag(group_size, tag_list, education_prob)

# Write file (const)
writer = pd.ExcelWriter('writerTest.xlsx', engine='xlsxwriter')
education_group.to_excel(writer, sheet_name="sheet 1", index=False)
art_group.to_excel(writer, sheet_name="sheet 2", index=False)
social_group.to_excel(writer, sheet_name="sheet 3", index=False)
writer.save()

print("Write end")

