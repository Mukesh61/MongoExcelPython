import pymongo
import xlrd
loc =("C:/Users/Mukesh/Desktop/pythonTest.xlsx")
wb=xlrd.open_workbook(loc)
sheet=wb.sheet_by_index(0)
svalue=sheet.cell_value(0,0)
svalue1=sheet.cell_value(0,1)
print(svalue)
wb.release_resources()
del wb

myclient = pymongo.MongoClient('mongodb://localhost:27017/')
mydb=myclient["mukesh"]
mycol=mydb["InfosysTest"]
x=mycol.find_one()
print(x)
newi={"name":svalue,"favorite super hero":svalue1}
myquery = { "name": svalue }
mydoc = mycol.find_one(myquery)
if(mydoc is None):
    mycol.insert_one(newi)
    print("i am not present")
else:
    print("hi baby i am there")
for y in mycol.find():
    print(y)




