import smtplib
from openpyxl import load_workbook
from datetime import datetime,date


workbook = load_workbook(filename=r"C:\Users\MDD\Desktop\python101\data_1.xlsx") #filepath
sheet = workbook.active

birthday_count=0
anniversary_count=0
kids_birthdays_count=0
birthdays={}
anniversaries={}
kids_birthdays={}

for value in sheet.iter_rows(min_row=2,min_col=1,values_only=True):	
	if(isinstance(value[2],datetime)):
		if(value[2].date().month==date.today().month):
			birthday_count+=1
			print(value[2].date(),value[0])
			birthdays.update({value[0]:value[2]})
print(f"This month has {birthday_count} birthdays,")
#print(birthdays)

	# birthday_count=0
	# if(isinstance(value[2],datetime)):
	# 	if(value[2].date().month==date.today().month):
	# 		birthday_count+=1
	# 		print(birthday_count)
	# 		print(value[2].date())
	# 		print(f"This month has {birthday_count} birthdays")	
	# 	else:
	# 		print(value)
		# if(value[l].date().month==date.today().month):
		# 	print(value[l].date())
	# ifwhile i<len(value):
	# 	print(value[i])
	# 	i+=1







	# l=0
	# list={}
	# while l<len(value):
	# 	if(isinstance(value[l],datetime)):
	# 		print(value[l])
	# 		if(value[l].date().month==date.today().month):
	# 			print(l)
	# 			print(value)
	# 			print(f"This month has birthdays on {value[l].date()}")
	# 	l+=1
				#print(f"This month birthdays for {value[0]} on {value[2].date()}")

	#if(value[3] is not None and isinstance(value[3],datetime)):
	#	if(value[3].date().month==date.today().month):
	#		print(f"This month anniversary for {value[0]} on {value[3].date()}")
	#if(value[4] is not None and isinstance(value[4],datetime)):
	#	if(value[4].date().month==date.today().month):
	#		print(f"This month kids birthdays for {value[0]} on {value[4].date()}")
			


