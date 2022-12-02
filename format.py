import os.path
from os import path
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from datetime import date

#Function to store all folders in a list
def list_paths(path):
    directories = [x[1] for x in os.walk(path)]
    non_empty_dirs = [x for x in directories if x] # filter out empty lists
    return [item for subitem in non_empty_dirs for item in subitem] # flatten the list

#function to edit dataframe
def edit_dataframe(foldersList,data,pathToDataset):
    for folder in foldersList:
            if folder.split(' ',1)[0]==data['Number'][row] or (folder.split(' ',1)[0]+' ' + folder.split(' ',2)[1])==data['Number'][row]:
                currentDir=os.path.join(pathToDataset,folder)
                for filename in os.listdir(currentDir):
                    if filename=="internal email.msg":
                        data.at[row,"Internal email"]= 'Yes'
                    if filename=="outside email.msg":
                        data.at[row,"Outside email"]= 'Yes'
                    if filename=="Number email.msg":
                        data.at[row,"Number email"]= 'Yes'
                    if filename[0].isdigit(): #to find Original date
                        currentFile= os.path.join(currentDir, filename)
                        doc_zip = zipfile.ZipFile(currentFile)
                        doc_xml = doc_zip.read("word/document.xml")
                        root = ET.fromstring(doc_xml)
                        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                        text_elements = root.findall(".//w:t", namespace)
                        dateString=""
                        originalDate=""
                        for i in range(len(text_elements)):
                            dateString += text_elements[i].text
                        # for c in dateString:

                        # print(dateString)
                        print("\n")
                break

# Load the xlsx file
pathToExcel="C://Users//Daryl//Documents//Coding_projects//Python//testing//python test//tracker.xlsx"
excel_data = pd.read_excel(pathToExcel)
data = pd.DataFrame(excel_data, columns=['Number', 'PIC', 'Original Date','Internal email', 'Outside email','Number email', 'Upload','Upload date', 'Today date', 'Comments'])

#Load dataset
pathToDataset = "C://Users//Daryl//Documents//Coding_projects//Python//testing//python test//dataset"
folders=os.listdir(pathToDataset)

#Store all folders in a list, then format the names according to excel 
foldersList=list_paths(pathToDataset)

#Editing Excel File
for row in data.index:
    if len(str(data['Original Date'][row]))>3: #To check if original date is empty, 3 because empty means 'NAN' 
        continue
    else:
        today = date.today()
        data.at[row,"Today date"]=today
        edit_dataframe(foldersList,data,pathToDataset)

data.to_excel('UpdatedTracker.xlsx', index=False)
print("You have successfully updated the tracker! ")

            
