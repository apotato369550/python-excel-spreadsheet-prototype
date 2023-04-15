import openpyxl
from datetime import datetime
import os
import tkinter as tk
from tkinter import *
from tkinter import messagebox as mb
import itertools

class ARARMaker(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        self.max_width = 900
        self.max_height = 575

        self.title("ARAR Maker V0.1")
        self.geometry(str(self.max_width) + "x" + str(self.max_height))
        self.resizable(False, False)

        self.invoices = {}

        self.make_user_interface()
        self.create_workbook()
        self.load_invoices()
        self.mainloop()

        return

    def update_invoices(self):
        return
    
    def load_invoices(self):
        # load invoices into dictionary
        print("loading invoices")
        self.workbook = openpyxl.load_workbook(filename="ARAR.xlsx")
        self.worksheet = self.workbook.active
        for row in itertools.islice(self.worksheet.iter_rows(values_only=True), 1, None):            
            try:
                self.customer_name = row[0]
                self.invoice_number = int(row[1])
                self.amount_due = float(row[2])

                if not isinstance(row[3], datetime):
                    self.due_date = datetime.strptime(row[3], "%m/%d/%Y")  
                else:
                    self.due_date = row[3]

                new_entry = {
                    "customer_name": self.customer_name,
                    "amount": self.amount_due,
                    "due_date": self.due_date
                }

                print(new_entry)

                if self.invoice_number not in self.invoices:
                    self.invoices[self.invoice_number] = {}  # Removed unnecessary dictionary initialization
                self.invoices[self.invoice_number] = new_entry

            except IndexError as e:  # More specific exception handling
                print(e)
                mb.showerror("Load Invoices: Error Loading Invoices", "An error occurred while loading the invoices. Please check the 'ARAR.xlsx' file and make sure everything is in its proper format.")
                return
            print(self.invoices)
            


        # iterate through each row of excel spreadsheet
        # for each row, add each invoice to a dictionary
        return
        

    def create_workbook(self):
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "ARAR"

        if not os.path.exists("ARAR.xlsx"):
            # List of column headers
            self.headers = ["Customer Name", "Invoice Number", "Invoice Amount", "Due Date", "Days Overdue", "<0 days", "0-30 days", "31-60 days", "61-90 days", "91-120 days", "Over 120 days"]

            # Loop through the headers and set the values in the cells
            for index, header in enumerate(self.headers):
                self.worksheet.cell(row=1, column=index+1, value=header)

            self.file_name = "ARAR.xlsx"
            self.workbook.save(self.file_name)

    def add_invoice(self):
        try:
            self.invoice_number = int(self.invoice_number_entry_1.get())
        except:
            mb.showerror("Add invoice: Invalid Invoice Number", "Please enter a valid invoice number. Omit all nondigit characters in the input field")
            return
        
        if self.invoice_number <= 0:
            mb.showerror("Add invoice: Invoice number is equal to 0", "Please enter a valid invoice number greater than 0. The value entered must be greater than 0")
            return
        
        self.customer_name = self.name_entry_1.get()
        if not self.customer_name:
            mb.showerror("Add invoice: Customer Name is Empty", "Please enter a valid customer name. The input field must not be empty.")
            return
        
        try:
            self.due_date = datetime.strptime(self.date_entry_1.get(), "%m/%d/%Y")
            self.days_overdue = (self.due_date - datetime.today()).days
        except ValueError:
            mb.showerror("Add invoice: Invalid Due Date", "Please enter a valid date before today following the format: mm/dd/yyyy")
            return

        try:
            self.amount_due = float(self.amount_entry_1.get())
        except ValueError:
            mb.showerror("Add invoice: Invalid Amount Due", "Please enter a valid amount. Omit all nondigit characters in the input field")
            return

        if self.amount_due <= 0:
            mb.showerror("Add invoice: Amount due is equal to 0", "Please enter an amount greater than 0. The value entered must be greater than 0")
            return

        # load invoices

        # perform input validation:
            # nothing is blank
            # check if invoice number is int and positive
            # check if due date is in valid format
            # check if due date is not before current date
            # check if amount is float and positive
        # check if invoice exists
        # open a popup window to confirm if user wants to add invoice
            # if user confirms, add that invoice to current invoice spreadsheet
            # regardless, clear entries afterwards
        return

    def delete_invoice(self):
        # perform input validation:
            # nothing is blank
            # check if invoice number is int and positive
        # check if invoice exists in ARAR
            # if invoice number is the only input: use that criteria only
        # open a popup window to confirm if user wants to delete invoice 
            # display all valid data about that invoice
            # if user confirms, delete that invoice 
            # regardless, clear entries afterwards
        return
    
    def view_invoices(self):
        return

    def make_user_interface(self):
        self.user_interface_frame = Frame(self, width=self.max_width, height=self.max_height)
        self.user_interface_frame.pack()

        self.title_label = Label(self.user_interface_frame, text="ARAR Maker", font=("Courier", 30, "bold"))
        self.title_label.grid(row=0, column=1, pady=5, columnspan=2)
        
        # add invoice button

        self.add_invoice_label = Label(self.user_interface_frame, text="Add an invoice", font=("Arial", 20, "bold"))
        self.add_invoice_label.grid(row=1, column=0, pady=5, columnspan=2)

        self.invoice_number_label_1 = Label(self.user_interface_frame, text="Invoice #")
        self.invoice_number_label_1.grid(row=2, column=0, pady=5)
        self.invoice_number_entry_1 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.invoice_number_entry_1.grid(row=3, column=0, pady=5, padx=7)

        self.name_label_1 = Label(self.user_interface_frame, text="Customer Name")
        self.name_label_1.grid(row=2, column=1, pady=5)
        self.name_entry_1 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.name_entry_1.grid(row=3, column=1, pady=5, padx=7)

        self.date_label_1 = Label(self.user_interface_frame, text="Due Date (mm/dd/yyyy)",)
        self.date_label_1.grid(row=2, column=2, pady=5)
        self.date_entry_1 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.date_entry_1.grid(row=3, column=2, pady=5, padx=7)

        self.amount_label_1 = Label(self.user_interface_frame, text="Amount",)
        self.amount_label_1.grid(row=2, column=3, pady=5)
        self.amount_entry_1 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.amount_entry_1.grid(row=3, column=3, pady=5, padx=7)

        # put command here
        self.add_invoice_button = Button(self.user_interface_frame, text="Add Invoice", font=("Courier", 15, "bold"))
        self.add_invoice_button.grid(row=1, column=2, columnspan=2, pady=5)

        # delete invoice feature

        self.delete_invoice_label = Label(self.user_interface_frame, text="Delete an Invoice", font=("Arial", 20, "bold"))
        self.delete_invoice_label.grid(row=5, column=0, pady=5, columnspan=2)
        
        self.invoice_number_label_2 = Label(self.user_interface_frame, text="Invoice #")
        self.invoice_number_label_2.grid(row=6, column=0, pady=5)
        self.invoice_number_entry_2 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.invoice_number_entry_2.grid(row=7, column=0, pady=5, padx=7)

        self.name_label_2 = Label(self.user_interface_frame, text="Customer Name (Optional)")
        self.name_label_2.grid(row=6, column=1, pady=5)
        self.name_entry_2 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.name_entry_2.grid(row=7, column=1, pady=5, padx=7)

        self.date_label_2 = Label(self.user_interface_frame, text="Due Date (mm/dd/yyyy) (Optional)",)
        self.date_label_2.grid(row=6, column=2, pady=5)
        self.date_entry_2 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.date_entry_2.grid(row=7, column=2, pady=5, padx=7)

        self.amount_label_2 = Label(self.user_interface_frame, text="Amount (Optional)",)
        self.amount_label_2.grid(row=6, column=3, pady=5)
        self.amount_entry_2 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.amount_entry_2.grid(row=7, column=3, pady=5, padx=7)

        self.delete_invoice_button = Button(self.user_interface_frame, text="Delete Invoice", font=("Courier", 15, "bold"))
        self.delete_invoice_button.grid(row=5, column=2, columnspan=2, pady=5)

        # listbox and view invoices button
        
        self.item_listbox_label = Label(self.user_interface_frame, text="AR Aging Report Overview", font=("Arial", 13))
        self.item_listbox_label.grid(row=9, column=1, columnspan=2, pady=5)

        self.item_listbox = Listbox(self.user_interface_frame, width=65, height=12, selectmode="SINGLE", borderwidth=5)
        self.item_listbox.grid(row=10, column=1, columnspan=2, pady=5)

        self.view_invoices_button = Button(self.user_interface_frame, text="View AR Aging Report", font=("Courier", 15, "bold"))
        self.view_invoices_button.grid(row=11, column=1, columnspan=2, pady=5)



if __name__ == "__main__":
    arar_maker = ARARMaker()