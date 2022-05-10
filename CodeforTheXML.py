import xlrd
#import a libart that will read data from a excel sheet
import xlwt
from xlwt import Workbook
import fileinput
from shutil import copyfile

#path to the excel file
loc = (r"localpath")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
wb = Workbook()


index=1
while(index<250):
    #opening the file and reading it and puting it all in the variable 'data'
    with open(r'xml.txt','r') as file:
        data=file.read()
       

        #the value of the cell being stored here to compare 
        Id =(sheet.cell_value(index,1))

        #checking the type of the variable Id if it is a float it will be changed to a string if it is a string then it will be given the value '#N/A'
        if type(Id) is float:
            Id = int(sheet.cell_value(index,1))
        else:
            Id = ("#N/A")

        #the text that needs to be replaced
        originalText = "<DefaultValue></DefaultValue>"

        #the replacement text will given the value from the excel sheet
        #***THE VALUE BEING ADDED HAS TO BE A STRING*****
        replacementText = "<DefaultValue>" + str(Id) + "</DefaultValue>"

        #replacing two strings 
        data=data.replace(originalText,replacementText)

        #creating the variable the new file will be given the name from the excel sheet
        newNameOfFile=sheet.cell_value(index,2)+'.txt'
    #creating the new file
    f=open((newNameOfFile),'w')
    f.write(data)

    #updating the index so the while loop will not be infinite
    index+=1

print ("Done")
