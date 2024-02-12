import xlsxwriter
import pandas as pd
import numpy as np

seed_value = 69
np.random.seed(seed_value)

# Constants
file_name = 'Student_data_2_2566.xlsx'
sheet_numbers = [13, 14, 15, 16]
tag_list = [
    "วิชาการ", "วิชาชีพ", "กิจการนักศึกษา", "เพิ่มทักษะ", "เทคโนโลยี",
    "บันเทิง", "ศิลปะ", "การแสดง", "ศาสนา","เทศกาล", 
    "สุขภาพ", "เข้าสังคม", "จิตอาสา", "ท่องเที่ยว", "กีฬา"
]
tag_probabilities = {
    "education": [0.12, 0.12, 0.12, 0.12, 0.12, 
                  0.02, 0.06, 0.02, 0.06, 0.06,
                   0.06, 0.06, 0.02, 0.02, 0.02],
    "art": [0.06, 0.06, 0.06, 0.06, 0.02, 
            0.12, 0.12, 0.12, 0.12, 0.12,
            0.02, 0.06, 0.02, 0.02, 0.02],
    "social": [0.06, 0.02, 0.06, 0.06, 0.02, 
               0.06, 0.02, 0.02, 0.02, 0.06,
               0.12, 0.12, 0.12, 0.12, 0.12]
}

# Read and concatenate data frames
frames = [pd.read_excel(file_name, sheet_name=num) for num in sheet_numbers]
merged_frames = pd.concat(frames).sample(frac=1).reset_index(drop=True)

# Split data
group_size = len(merged_frames) // 3
education_group, art_group, social_group = np.array_split(merged_frames, 3)


def generate_tags(size, tag_list, probabilities):
    tags = []
    for _ in range(size):
        rnd_tags = np.random.choice(tag_list, 3, p=probabilities, replace=False)
        tags.append(rnd_tags)
    return tags

# Generate tags for each group
education_tags = generate_tags(group_size, tag_list, tag_probabilities["education"])
art_tags = generate_tags(group_size, tag_list, tag_probabilities["art"])
social_tags = generate_tags(group_size, tag_list, tag_probabilities["social"])

# Create data frames for tags
education_tags_df = pd.DataFrame(education_tags, columns=['ความสนใจ 1', 'ความสนใจ 2', 'ความสนใจ 3'])
art_tags_df = pd.DataFrame(art_tags, columns=['ความสนใจ 1', 'ความสนใจ 2', 'ความสนใจ 3'])
social_tags_df = pd.DataFrame(social_tags, columns=['ความสนใจ 1', 'ความสนใจ 2', 'ความสนใจ 3'])

# Merge groups with tags
education_group = pd.concat([education_group.reset_index(drop=True), education_tags_df], axis=1)
art_group = pd.concat([art_group.reset_index(drop=True), art_tags_df], axis=1)
social_group = pd.concat([social_group.reset_index(drop=True), social_tags_df], axis=1)

# Write to Excel file
with pd.ExcelWriter('Generated_data.xlsx', engine='xlsxwriter') as writer:
    education_group.to_excel(writer, sheet_name="Education", index=False)
    art_group.to_excel(writer, sheet_name="Art", index=False)
    social_group.to_excel(writer, sheet_name="Social", index=False)

print("Write end")
