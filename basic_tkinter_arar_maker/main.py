import openpyxl
from datetime import datetime
import os
import tkinter as tk
from tkinter import *

class ARARMaker(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        self.max_width = 900
        self.max_height = 575

        self.title("ARAR Maker V0.1")
        self.geometry(str(self.max_width) + "x" + str(self.max_height))
        self.resizable(False, False)

        self.items = []

        self.make_user_interface()
        self.mainloop()

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