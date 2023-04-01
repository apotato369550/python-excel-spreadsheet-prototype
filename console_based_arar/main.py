import openpyxl
from datetime import datetime
from os import system

# create workbook and worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = "AR Aging Report"

# header/title rwo
worksheet['A1'] = "Customer Name"
worksheet['B1'] = "Invoice Number"
worksheet['C1'] = "Invoice Amount"
worksheet['D1'] = "Due Date"
worksheet['E1'] = "Days Overdue"

invoices = {}
invoice_totals = 0 
period_totals = [0, 0, 0, 0, 0]

print("Welcome user to console-based ARAR-maker V1!")

filename = input("To begin, please enter the filename you wish to save the report as (entering no name will make 'arar.xlsx' the default filename): ")

if filename:
    filename = f"{filename}.xlsx"
else:
    filename = "arar.xlsx"

while True:
    customer_name = input("Please enter the name of the customer: ")
    
    while True:
        try:
            amount_due = input("Please enter amount due by customer: ")
            float(amount_due)
            break
        except ValueError:
            print('Please enter a valid amount.')



    while True:
        try:
            due_date = datetime.strptime(input("Please enter the payement's due date (format: mm/dd/yyyy)"), "%m/%d/%Y")
            break
        except ValueError:
            print("Please enter a valid date in the following format: mm/dd/yyyy")
    
    break

print(f"Customer's name: {customer_name}")
print(f"Amount Due: {str(amount_due)}")
print(f"Date Due: {due_date}")


workbook.save(f"{filename}")

while True:
    view = input("Do you wish to run Excel and view the created file? (Y/N): ").lower()
    if view == 'y':
        system(f"start excel.exe {filename}")
        break
    elif view == 'n':
        print("See you soon!")
        break
    else:
        print("Please enter a valid response.")


