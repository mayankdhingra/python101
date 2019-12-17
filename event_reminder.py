import smtplib
from openpyxl import load_workbook
from datetime import datetime,date


workbook = load_workbook(filename=r"C:\Users\MDD\Desktop\python101\data_1.xlsx") #filepath
#workbook = load_workbook(filename=r"C:\Users\Shuttl\Desktop\Python\python101\data.xlsx")
sheet = workbook.active

for value in sheet.iter_rows(min_row=2,min_col=1,values_only=True):
	if(value[2] is not None):
		if(value[2].date().month==date.today().month):
			print(f"same month birthdays for {value[0]} on {value[2].date()}")

	if(value[3] is not None):
		if(value[3].date().month==date.today().month):
			print(f"same month anniversary for {value[0]} on {value[3].date()}")
	if(value[4] is not None):
		if(value[4].date().month==date.today().month):
			print(f"same month kids birthdays for {value[0]} on {value[4].date()}")
			


