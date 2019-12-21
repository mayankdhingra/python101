import smtplib,re,operator,sys
from openpyxl import load_workbook
from datetime import datetime,date
from io import StringIO


workbook = load_workbook(filename=r"C:\Users\MDD\Desktop\python101\data_1.xlsx") #filepath

sheet = workbook.active

birthday_count=0
anniversary_count=0
kids_birthdays_count=0
birthdays={}
anniversaries={}
kids_birthdays={}

old_stdout = sys.stdout # Memorize the default stdout stream
result= StringIO()

sys.stdout = result

def change_date_format(dt):
        return re.sub(r'(\d{4})-(\d{1,2})-(\d{1,2})', '\\3-\\2-\\1', dt)

for value in sheet.iter_rows(min_row=2,min_col=1,values_only=True):	
	if(isinstance(value[2],datetime)):
		if(value[2].date().month==date.today().month):
			birthday_count+=1
			birthdays.update({value[0]:value[2]})
	

	if(isinstance(value[3],datetime)):
		if(value[3].date().month==date.today().month):
			anniversary_count+=1
			anniversaries.update({value[0]:value[3]})

	if(isinstance(value[4],datetime)):
		if(value[4].date().month==date.today().month):
			kids_birthdays_count+=1
			kids_birthdays.update({value[0]:value[4]})


birthdays=dict(sorted(birthdays.items(), key=operator.itemgetter(0)))
print(f"This month has {birthday_count} birthdays,")
for key in birthdays:
	print(key,'->',change_date_format(str(birthdays[key].date())))
	b_message = str(key) + "->" 
print("\n")
anniversaries=dict(sorted(anniversaries.items(), key=operator.itemgetter(0)))
print(f"This month has {anniversary_count} anniversaries,")
for key in anniversaries:
	print(key,'->',change_date_format(str(anniversaries[key].date())))
print("\n")
kids_birthdays=dict(sorted(kids_birthdays.items(), key=operator.itemgetter(0)))
print(f"This month has {kids_birthdays_count} kids_birthdays,")
for key in kids_birthdays:
	print(key,'->',change_date_format(str(kids_birthdays[key].date())))
print("\n")

sys.stdout = old_stdout # Put the old stream back in place
whatWasPrinted = "\n" + result.getvalue() # Return a str containing the entire contents of the buffer.

print(f"message is {whatWasPrinted}")

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login("email id", "password") #email id and password
server.sendmail("email id","receipient email id",whatWasPrinted)

