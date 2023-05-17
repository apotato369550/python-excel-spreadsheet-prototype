import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
from customtkinter import *
import customtkinter as ctk
from tkinter import messagebox as mb

class POUI(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")  
        self.max_width = 600
        self.max_height = 800

        self.title("InvoiceMate Version 1.0: Purchase Order Generator")
        self.geometry(str(self.max_width) + "x" + str(self.max_height))
        self.resizable(False, False)


        self.make_user_interface()
    
    def make_user_interface(self):
        self.user_interface_frame = CTkFrame(self, width=self.max_width, height=self.max_height)
        self.user_interface_frame.pack()

        self.po_maker_label = CTkLabel(self.user_interface_frame, text="Receipt Maker", font=("Courier", 40, "bold"))
        self.po_maker_label.grid(row=0, column=0, pady=5, columnspan=2)

        self.basic_info_label = CTkLabel(self.user_interface_frame, text="Basic Info", font=("Arial", 25, "bold"))
        self.basic_info_label.grid(row=1, column=0, pady=5)

        self.invoice_number_label = CTkLabel(self.user_interface_frame, text="Invoice Number: ")
        self.invoice_number_entry = CTkEntry(self.user_interface_frame, width=275)
        self.invoice_number_label.grid(row=2, column=0)
        self.invoice_number_entry.grid(row=3, column=0)

        self.date_label = CTkLabel(self.user_interface_frame, text="SC/PWD Discount Rate (%) (in decimal): ")
        self.date_entry = CTkEntry(self.user_interface_frame, width=275)
        self.date_label.grid(row=2, column=1)
        self.date_entry.grid(row=3, column=1)

        self.customer_name_label = CTkLabel(self.user_interface_frame, text="Customer Name: ")
        self.customer_name_entry = CTkEntry(self.user_interface_frame, width=275)
        self.customer_name_label.grid(row=4, column=0, padx=7)
        self.customer_name_entry.grid(row=5, column=0, padx=7)

        self.tin_label = CTkLabel(self.user_interface_frame, text="TIN: ")
        self.tin_entry = CTkEntry(self.user_interface_frame, width=275)
        self.tin_label.grid(row=4, column=1, padx=7)
        self.tin_entry.grid(row=5, column=1, padx=7)

        self.terms_label = CTkLabel(self.user_interface_frame, text="Terms: ")
        self.terms_entry = CTkEntry(self.user_interface_frame, width=275)
        self.terms_label.grid(row=6, column=0, padx=7)
        self.terms_entry.grid(row=7, column=0, padx=7)

        self.osca_pwd_id_label = CTkLabel(self.user_interface_frame, text="OSCA/PWD ID (if applicable): ")
        self.osca_pwd_id_entry = CTkEntry(self.user_interface_frame, width=275)
        self.osca_pwd_id_label.grid(row=6, column=1, padx=7)
        self.osca_pwd_id_entry.grid(row=7, column=1, padx=7)

        self.address_label = CTkLabel(self.user_interface_frame, text="Address: ")
        self.address_entry = CTkEntry(self.user_interface_frame, width=275)
        self.address_label.grid(row=8, column=0, padx=7)
        self.address_entry.grid(row=9, column=0, padx=7)

        self.business_style_label = CTkLabel(self.user_interface_frame, text="Business Style: ")
        self.business_style_entry = CTkEntry(self.user_interface_frame, width=275)
        self.business_style_label.grid(row=8, column=1, padx=7)
        self.business_style_entry.grid(row=9, column=1, padx=7)


        self.sc_pwd_discount_label = CTkLabel(self.user_interface_frame, text="SC/PWD Discount Rate (%) (in decimal): ")
        self.sc_pwd_discount_entry = CTkEntry(self.user_interface_frame, width=275)
        self.sc_pwd_discount_label.grid(row=10, column=0)
        self.sc_pwd_discount_entry.grid(row=11, column=0)

        self.withholding_tax_label = CTkLabel(self.user_interface_frame, text="Withholding Tax Rate (%) (in decimal): ")
        self.withholding_tax_entry = CTkEntry(self.user_interface_frame, width=275)
        self.withholding_tax_label.grid(row=10, column=1)
        self.withholding_tax_entry.grid(row=11, column=1)

        # add/remove items
        self.basic_info_label = CTkLabel(self.user_interface_frame, text="Add/Remove Items", font=("Arial", 25, "bold"))
        self.basic_info_label.grid(row=12, column=0, pady=5)

        self.quantity_label = CTkLabel(self.user_interface_frame, text="Quantity: ")
        self.quantity_entry = CTkEntry(self.user_interface_frame, width=275)
        self.quantity_label.grid(row=13, column=0, pady=5)
        self.quantity_entry.grid(row=14, column=0, pady=5)

        self.unit_price_label = CTkLabel(self.user_interface_frame, text="Unit Price: ")
        self.unit_price_entry = CTkEntry(self.user_interface_frame, width=275)
        self.unit_price_label.grid(row=15, column=0, padx=5)
        self.unit_price_entry.grid(row=16, column=0, padx=5)

        self.item_description_label = CTkLabel(self.user_interface_frame, text="Item Description", font=("Arial", 13))
        self.item_description_label.grid(row=17, column=0, pady=5)

        self.item_textbox = CTkTextbox(self.user_interface_frame, width=270, height=195)
        self.item_textbox.grid(row=18, column=0, pady=5, rowspan=8)

        self.unit_label = CTkLabel(self.user_interface_frame, text="Unit (pcs., kgs, boxes, etc.): ")
        self.unit_entry = CTkEntry(self.user_interface_frame, width=275)
        self.unit_label.grid(row=13, column=1, padx=5)
        self.unit_entry.grid(row=14, column=1, padx=5)

        self.article_label = CTkLabel(self.user_interface_frame, text="Article: ")
        self.article_entry = CTkEntry(self.user_interface_frame, width=275)
        self.article_label.grid(row=15, column=1, padx=5)
        self.article_entry.grid(row=16, column=1, padx=5)

        self.item_textbox_label = CTkLabel(self.user_interface_frame, text="Item List", font=("Arial", 13))
        self.item_textbox_label.grid(row=17, column=1, pady=5)

        # edit color
        self.item_selection = CTkOptionMenu(self.user_interface_frame, values=["Option 1", "Option 2", "Option 3"], fg_color="#343638", button_color="#343638", button_hover_color="#565b5e")
        self.item_selection.grid(row=18, column=1, pady=5)

        self.add_item_button = CTkButton(self.user_interface_frame, text="Add Item", fg_color="#343638", border_color="#565b5e", width=275)
        self.add_item_button.grid(row=19, column=1, pady=5)

        self.remove_item_button = CTkButton(self.user_interface_frame, text="Remove Item", fg_color="#343638", border_color="#565b5e", width=275)
        self.remove_item_button.grid(row=20, column=1, pady=5)
        

        self.record_receipt_button = CTkButton(self.user_interface_frame, text="Record Receipt", fg_color="#343638", border_color="#565b5e", width=275)
        self.record_receipt_button.grid(row=22, column=1, pady=5)

        def checkbox_event():
            print("checkbox toggled, current value:", self.check_var.get())

        self.check_var = StringVar(value="on")

        self.arar_checkbox = CTkCheckBox(self.user_interface_frame, text="Add Current Invoice to AR Aging Report", command=checkbox_event, variable=self.check_var, onvalue="on", offvalue="off", hover_color="#565b5e", fg_color="#343638",)
        self.arar_checkbox.grid(row=21, column=1, pady=5)

        return
