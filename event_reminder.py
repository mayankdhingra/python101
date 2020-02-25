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
this_weeks_birthday_count = 0
this_weeks_anniversary_count=0
this_weeks_kids_birthdays_count=0
this_weeks_birthdays={}
this_weeks_anniversaries={}
this_weeks_kids_birthdays={}
next_month_birthday_count = 0
next_month_anniversary_count=0
next_month_kids_birthdays_count=0
next_month_birthdays={}
next_month_anniversaries={}
next_month_kids_birthdays={}


sender_email = "a@gmail.com"  # Enter your address
receiver_email = "d@gmail.com"  # Enter receiver address


old_stdout = sys.stdout # Memorize the default stdout stream
result= StringIO()

sys.stdout = result

def change_date_format(dt):
        return re.sub(r'(\d{4})-(\d{1,2})-(\d{1,2})', '\\3-\\2-\\1', dt)

def is_event_this_week(date_today,date_event):
	day_since= (date_event-date_today).days
	day_today=date_today.weekday()
	map_this_week={0:[0,6],1:[-1,5],2:[-2,4],3:[-3,3],4:[-4,2],5:[-5,1],6:[-6,0]}
	range=map_this_week[day_today]
	if(day_since>=range[0] and day_since<=range[1]):
		return True
	else:
		return False
	

for value in sheet.iter_rows(min_row=2,min_col=1,values_only=True):	
	if(isinstance(value[2],datetime)):
		if(value[2].date().month==date.today().month):
			event_this_week_or_not=is_event_this_week(date.today(),value[2].date())
			if event_this_week_or_not:
				this_weeks_birthday_count+=1
				#this_weeks_birthdays.update({value[0]:value[2]})
			birthday_count+=1
			birthdays.update({value[0]:value[2]})
		if(date.today().month==12):
			if(value[2].date().month==1):
				next_month_birthday_count+=1
				next_month_birthdays.update({value[0]:value[2]})
		else:
			#print("here")
			if(value[2].date().month==date.today().month+1):
				next_month_birthday_count+=1
				next_month_birthdays.update({value[0]:value[2]})


	if(isinstance(value[3],datetime)):
		if(value[3].date().month==date.today().month):
			anniversary_count+=1
			anniversaries.update({value[0]:value[3]})

		if(date.today().month==12):
			if(value[3].date().month==1):
				next_month_birthday_count+=1
				next_month_birthdays.update({value[0]:value[3]})
		else:
			if(value[3].date().month==date.today().month+1):
				next_month_birthday_count+=1
				next_month_birthdays.update({value[0]:value[3]})

	if(isinstance(value[4],datetime)):
		if(value[4].date().month==date.today().month):
			kids_birthdays_count+=1
			kids_birthdays.update({value[0]:value[4]})

		if(date.today().month==12):
			if(value[4].date().month==1):
				next_month_birthday_count+=1
				next_month_birthdays.update({value[0]:value[4]})
		else:
			#print("here")
			if(value[4].date().month==date.today().month+1):
				next_month_birthday_count+=1
				next_month_birthdays.update({value[0]:value[4]})

#print(f"This Month has {birthday_count} birthdays, {anniversary_count} anniversaries & {kids_birthdays_count} kids birthdays \n")
if this_weeks_birthday_count:
	print(f"This week has {this_weeks_birthday_count} birthdays \n")
	
birthdays=dict(sorted(birthdays.items(), key=operator.itemgetter(1)))
print(f"This month has {birthday_count} birthdays: \n")
for key in birthdays:
	print(key,'->',change_date_format(str(birthdays[key].date())))
	b_message = str(key) + "->" 
print("\n")

anniversaries=dict(sorted(anniversaries.items(), key=operator.itemgetter(1)))
if anniversary_count:
	print(f"This month has {anniversary_count} anniversaries: \n")
else:
	print(f"This month has No anniversaries: \n")
	
for key in anniversaries:
	print(key,'->',change_date_format(str(anniversaries[key].date())))
print("\n")

kids_birthdays=dict(sorted(kids_birthdays.items(), key=operator.itemgetter(1)))
print(f"This month has {kids_birthdays_count} kids birthdays: \n")
for key in kids_birthdays:
	print(key,'->',change_date_format(str(kids_birthdays[key].date())))
print("\n")

print(f"Next Month has {next_month_birthday_count} birthdays, {next_month_anniversary_count} anniversaries & {next_month_kids_birthdays_count} kids birthdays \n")

next_month_birthdays=dict(sorted(next_month_birthdays.items(), key=operator.itemgetter(1)))
print(f"Next month has {next_month_birthday_count} birthdays: \n")
for key in next_month_birthdays:
	print(key,'->',change_date_format(str(next_month_birthdays[key].date())))
	b_message = str(key) + "->" 
print("\n")

next_month_anniversaries=dict(sorted(next_month_anniversaries.items(), key=operator.itemgetter(1)))
print(f"Next month has {next_month_anniversary_count} anniversaries: \n")
for key in next_month_anniversaries:
	print(key,'->',change_date_format(str(next_month_anniversaries[key].date())))
print("\n")

next_month_kids_birthdays=dict(sorted(next_month_kids_birthdays.items(), key=operator.itemgetter(1)))
print(f"Next month has {next_month_kids_birthdays_count} kids birthdays: \n")
for key in next_month_kids_birthdays:
	print(key,'->',change_date_format(str(next_month_kids_birthdays[key].date())))
print("\n")

total_events = birthday_count + anniversary_count + kids_birthdays_count
next_month_total_events= next_month_birthday_count + next_month_anniversary_count + next_month_kids_birthdays_count

sys.stdout = old_stdout # Put the old stream back in place
message = """\
Subject: This Month's Events (Birthdays, Anniversaries etc): """ + str(total_events) + """


"""


whatWasPrinted = message + result.getvalue() # Return a str containing the entire contents of the buffer.

print(f"message is {whatWasPrinted}")

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
#server.login("email id", "password") #email id and password
#server.sendmail("email id","receipient email id",whatWasPrinted)
#server.sendmail(sender_email,receiver_email,whatWasPrinted)
