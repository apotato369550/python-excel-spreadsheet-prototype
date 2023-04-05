import openpyxl
from datetime import datetime
import os
import tkinter as tk
from tkinter import *

class ARARMaker(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        self.max_width = 900
        self.max_height = 700

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
        self.title_label.grid(row=0, column=0, pady=5, columnspan=2)
        
        self.add_invoice_label = Label(self.user_interface_frame, text="Add an invoice", font=("Arial", 20, "bold"))
        self.add_invoice_label.grid(row=1, column=0, pady=5, columnspan=2)

        self.invoice_number_label_1 = Label(self.user_interface_frame, text="Invoice #")
        self.invoice_number_label_1.grid(row=2, column=0, pady=5)
        self.invoice_number_entry_1 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.invoice_number_entry_1.grid(row=3, column=0)

        self.name_label_1 = Label(self.user_interface_frame, text="Customer Name")
        self.name_label_1.grid(row=2, column=1, pady=5)
        self.name_entry_1 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.name_entry_1.grid(row=3, column=1)

        self.date_label_1 = Label(self.user_interface_frame, text="Due Date (mm/dd/yyyy)",)
        self.date_label_1.grid(row=2, column=2, pady=5)
        self.date_entry_1 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.date_entry_1.grid(row=3, column=2)

        self.amount_label_1 = Label(self.user_interface_frame, text="Amount",)
        self.amount_label_1.grid(row=2, column=3, pady=5)
        self.amount_entry_1 = Entry(self.user_interface_frame, width=30, borderwidth=5)
        self.amount_entry_1.grid(row=3, column=3)

        self.delete_invoice_label = Label(self.user_interface_frame, text="Delete an Invoice", font=("Arial", 20, "bold"))
        self.delete_invoice_label.grid(row=7, column=0, pady=5, columnspan=2)


if __name__ == "__main__":
    arar_maker = ARARMaker()