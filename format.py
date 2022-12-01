import os.path
from os import path
import pandas as pd

# Load the xlsx file
pathToExcel="C://Users//Daryl//Documents//Coding_projects//Python//testing//python test//tracker.xlsx"
excel_data = pd.read_excel(pathToExcel)
data = pd.DataFrame(excel_data, columns=['Number', 'PIC', 'Original Date','Internal email', 'Outside email', 'Upload','Upload date', 'Today date', 'Comments'])

#Load dataset
pathToDataset = "C://Users//Daryl//Documents//Coding_projects//Python//testing//python test//dataset"
folders=os.listdir(pathToDataset)


#Editing file part
for row in data.index:
    if len(str(data['Original Date'][row]))>3: #To check if original date is empty, unsure why it is 3, suppose to be 0 but when I print the blank cell got length of 3
        continue
    else:
        print(data['Number'][row])






# for subdir, dirs, files in os.walk(pathToDataset):
#     print(files)
