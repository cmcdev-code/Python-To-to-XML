from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient import discovery
from pprint import pprint
import xlrd
import xlwt
from xlwt import Workbook
import fileinput
from shutil import copyfile

#The Excel file can not be type .xlsx
loc = (r"localPathOfExcelFile")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

wb = Workbook()
index=1
arrayOfTextfile=[]
while(index<250): #250 is the size of the rows in the excel file
    with open(r'NameOfTextFile.txt','r') as file:
        
        data=file.read()#reading entire .txt file and assigning it the variabel 'data'
        
        #setting the value of what is in the Excel cell at that position to 'Id'
        Id = (sheet.cell_value(index,1))
        
        #Checking the type of the data that is in the Excel cell
        if type(Id) is float:
            Id = int(sheet.cell_value(index,1))
        else:
            Id = ("")#Change it to the value that needs to be there
            
        #originalText is the value in the text file that will be replaced   
        originalText = "<DefaultValue></DefaultValue>"
        
        #replacementText is the text that will be replacing the the 'originalText'
        replacementText = "<DefaultValue>" + str(Id) + "</DefaultValue>"
       
        data=data.replace(originalText,replacementText)#actual swap of the text
        arrayOfTextfile.append(data) #adding the entire .txt file with the updated text to the 'arrayOfTextfile'
        index+=1 #updating the index

#downloaded .json keys to the Google Sheets api 
SERVICE_ACCOUNT_FILE = 'newKeys.json'

"""
  #https://www.googleapis.com/auth/spreadsheets.readonly , Allows read-only access to the user's sheets and their properties.
  #https://www.googleapis.com/auth/spreadsheets	         , Allows read/write access to the user's sheets and their properties.
  #https://www.googleapis.com/auth/drive.readonly	     , Allows read-only access to the user's file metadata and file content.
  #https://www.googleapis.com/auth/drive.file	         , Per-file access to files created or opened by the app.
  #https://www.googleapis.com/auth/drive	             , Full permissive scope to access all of a user's files. Request this scope only when it is strictly necessary.
"""
SCOPES = ['https://www.googleapis.com/auth/drive.readonly'] 

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

#https://docs.google.com/spreadsheets/d/Id Will Be here/edit#gid=000000000
SAMPLE_SPREADSHEET_ID = 'Id goes here'

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
#Read the sheet and print it out
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="NameOfSheet!A1:A251").execute()
values = result.get('values', [])
print(values)


ArrayOfArray=[arrayOfTextfile]
request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="NameOfSheet!A1", valueInputOption="RAW", body={"values":ArrayOfArray}).execute()#RAW	The values the user has entered will not be parsed and will be stored as-is.
                                                                                                                   #USER_ENTERED	The values will be parsed as if the user typed them into the UI. Numbers will stay as numbers, but strings may be converted to numbers, dates, etc. following the same rules that are applied when entering text into a cell via the Google Sheets UI.
#Google Sheet Updates 

