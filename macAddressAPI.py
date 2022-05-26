##################Libary that allows requests of the api data
from hashlib import new
import http.client
import requests
########Libaries that will allow the  porgram to read in data from an excel sheet
from operator import index
import xlwt
import xlrd
from xlwt import Workbook
import fileinput

##### Oppening the excel file and asigning it to a variable 
loc = (r"localPath")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
wb = Workbook()

index=0
array=[]
arrayOfJson=[]

##string that will change the values for the api request 
apiRequest="macaddress.io"##get the api request here 



#### Creating a loop that will read in the data and store it to the excel sheet
while(index<77):
    Id=(sheet.cell_value(index,3))
    array.append(Id)

    requestsString = apiRequest+array[index]
    print(requestsString)
    response = requests.get(requestsString)
    arrayOfJson.append(response.json())

    index+=1

####STUFFF SO IT CAN BE OUTPUTTED AS AN EXCEL FILE
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("jsonData")

header_font = xlwt.Font()
header_font.name = 'Arial'
header_font.bold = True

header_style = xlwt.XFStyle()
header_style.font = header_font
index=0
while(index<77): 
    sheet.write(index, 2, str(array[index]), header_style)
    sheet.write(index, 1, str(arrayOfJson[index]), header_style)
    index+=1
workbook.save('exceldata.xls')
