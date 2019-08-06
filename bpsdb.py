import openpyxl
import mysql.connector

mydb=mysql.connector.connect(
	host="localhost",
	user="root",
	password="fsociety",
	database="pysql"
)
mycursor=mydb.cursor()

sql="INSERT INTO `reciever`(`fname`, `lname`, `email`, `phone`, `address`, `town`, `businessname`) VALUES (%s,%s,%s,%s,%s,%s,%s)";

workbook=openpyxl.load_workbook('BPS.xlsx')
sheet=workbook.get_active_sheet()
for row in range(3,sheet.max_row+1):
	name=sheet['A'+str(row)].value
	allnames=str(name).split(" ")
	fname=""
	lname=""
	if len(allnames)>1:
		fname=allnames[0]
		lname=allnames[1]
	else:
		fname=allnames[0]
		lname='NULL'
	businessname=sheet['B'+str(row)].value
	contact=sheet['C'+str(row)].value
	if(contact==None):
		contact='NULL'
	else:
		contact='+254'+str(contact)
	location=sheet['D'+str(row)].value
	#print(name," RUns ",businessname,"Call ",contact,"IN ",location)
	print("INSERTING "+fname," ",lname,businessname,contact)
	val=(fname,lname,"null email",contact,'Null',location,businessname)

	mycursor.execute(sql,val)
	mydb.commit()
