import os.path
from os import path
import pandas as pd

#Function to store all folders in a lis
def list_paths(path):
    directories = [x[1] for x in os.walk(path)]
    non_empty_dirs = [x for x in directories if x] # filter out empty lists
    return [item for subitem in non_empty_dirs for item in subitem] # flatten the list

# Load the xlsx file
pathToExcel="C://Users//Daryl//Documents//Coding_projects//Python//testing//python test//tracker.xlsx"
excel_data = pd.read_excel(pathToExcel)
data = pd.DataFrame(excel_data, columns=['Number', 'PIC', 'Original Date','Internal email', 'Outside email','Number email', 'Upload','Upload date', 'Today date', 'Comments'])

#Load dataset
pathToDataset = "C://Users//Daryl//Documents//Coding_projects//Python//testing//python test//dataset"
folders=os.listdir(pathToDataset)

#Store all folders in a list, then format the names according to excel 
foldersList=list_paths(pathToDataset)
formattedFoldersList=[]
for folder in foldersList:
    formattedFoldersList.append(folder.split(' ',1)[0])


#Editing Excel File
for row in data.index:
    if len(str(data['Original Date'][row]))>3: #To check if original date is empty, 3 because empty means 'NAN' 
        continue
    else:
        for folder in foldersList:
            if folder.split(' ',1)[0]==data['Number'][row]:
                currentDir=os.path.join(pathToDataset,folder)
                for filename in os.listdir(currentDir):
                    if filename=="internal email.msg":
                        data.at[row,"Internal email"]= 'Yes'
                    if filename=="outside email.msg":
                        data.at[row,"Outside email"]= 'Yes'
                    if filename=="Number email.msg":
                        data.at[row,"Number email"]= 'Yes'
                if data['Internal email'][row] != 'Yes':
                    data.at[row,"Internal email"]= 'No'
                if data['Outside email'][row] != 'Yes':
                    data.at[row,"Outside email"]= 'No'
                if data['Number email'][row] != 'Yes':
                    data.at[row,"Number email"]= 'No'
                break

data.to_excel('UpdatedTracker.xlsx', index=False)
print("You have successfully updated the tracker! ")

            
