from docxtpl import DocxTemplate
from datetime import datetime
import os
from customtkinter import *
import customtkinter as ctk
from tkinter import messagebox as mb
import openpyxl
from openpyxl.utils import get_column_letter

class POUI(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")  
        self.max_width = 600
        self.max_height = 835

        self.title("InvoiceMate Version 1.0: Purchase Order Generator")
        self.geometry(str(self.max_width) + "x" + str(self.max_height))
        self.resizable(False, False)

        self.items = {}

        self.make_user_interface()


    def record_receipt(self):
        self.invoice_number = self.invoice_number_entry.get()
        self.customer_name = self.customer_name_entry.get()
        self.tin = self.tin_entry.get()
        self.terms = self.terms_entry.get()
        self.osca_pwd_id = self.osca_pwd_id_entry.get()
        self.address = self.address_entry.get()
        self.business_style = self.business_style_entry.get()
        self.withholding_tax = self.withholding_tax_entry.get()
        self.sc_pwd_discount = self.sc_pwd_discount_entry.get()
        self.date = self.date_entry.get()

        if (not self.invoice_number or not self.customer_name or not self.tin or not self.terms 
        or not self.osca_pwd_id or not self.address or not self.business_style):
            mb.showerror("Basic Info: Empty Fields", "Please make sure all fields are filled out. Please type 'N/A' in any of the fields that do not apply")
            return

        try:
            self.sc_pwd_discount = float(self.sc_pwd_discount)
            if self.sc_pwd_discount >= 1 and self.sc_pwd_discount < 0:
                mb.showerror("Basic Info: SC/PWD Discount Rate", "Please enter a valid discount rate in decimal. The value entered must be less than 1, but greater than 0.")
                return
        except:
            mb.showerror("Basic Info: SC/PWD Discount Rate", "Please enter a valid discount rate. Omit all nondigit characters in the input field.")
            return
        
        try:
            self.withholding_tax = float(self.withholding_tax)
            if self.withholding_tax >= 1 and self.withholding_tax < 0:
                mb.showerror("Basic Info: Withholding Tax Rate", "Please enter a valid withholding tax rate in decimal. The value entered must be less than 1, but greater than 0.")
                return
        except:
            mb.showerror("Basic Info: Withholding Tax Rate", "Please enter a valid withholding tax rate. Omit all nondigit characters in the input field.")
            return
        
        if not self.date:
            mb.showerror("Date Selector: Empty Fields", "Please select a valid date.")
            return

        if not bool(self.items):
            mb.showerror("Item Adder: No Items Added", "Please add at least (1) item to the receipt.")
            return

        try:
            self.date = datetime.strptime(self.date, "%m/%d/%Y")
            self.days_overdue = (self.date - datetime.today()).days
        except ValueError:
            mb.showerror("Basic Info: Invalid Date", "Please enter a valid date before today following the format: mm/dd/yyyy")
            return

        self.total_sales = 0
        self.sales_table_rows = []

        # fix bug
        for item_name, item_properties in self.items.items():
            print(item_name)
            print(item_properties)
            self.total_sales += item_properties["amount"]
            self.row = {}
            self.row["quantity"] = str(item_properties["quantity"])
            self.row["unit"] = item_properties["unit"]
            self.row["articles"] = item_name
            self.row["unit_price"] = str(item_properties["unit_price"])
            self.row["amount"] = str(item_properties["amount"])
            self.sales_table_rows.append(self.row)

        self.sc_pwd_discount = self.total_sales * self.sc_pwd_discount
        self.total_amount = self.total_sales - self.sc_pwd_discount
        self.withholding_tax = self.total_amount * self.withholding_tax
        self.total_amount_due = self.total_amount - self.withholding_tax

        self.doc = DocxTemplate("templates/po_template.docx")

        self.context = {
            "invoice_number": self.invoice_number,
            "sold_to": self.customer_name,
            "date": self.date,
            "tin": self.tin,
            "terms": self.terms,
            "address": self.address,
            "osca_pwd_id": self.osca_pwd_id,
            "business_style": self.business_style,
            "sales_table_rows": self.sales_table_rows,
            "total_sales": str(self.total_sales),
            "sc_pwd_discount": str(self.sc_pwd_discount),
            "total_amount": self.total_amount,
            "withholding_tax": self.withholding_tax,
            "total_amount_due": self.total_amount_due
        }

        self.doc.render(self.context)
        self.doc.save("receipts/invoice_number_" + self.invoice_number + ".docx")

        self.items = {}
        self.item_textbox.configure(state="normal")
        self.item_textbox.delete("1.0", "end")
        self.item_textbox.configure(state="disabled")

            # add to arar
        
        mb.showinfo("Receipt Successfully Created", "Receipt successfully created. Please check the '/output' folder for the finished receipt.")

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

        self.date_label = CTkLabel(self.user_interface_frame, text="Date (mm/dd/yy): ")
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

        self.item_textbox = CTkTextbox(self.user_interface_frame, width=270, height=215, font=("Arial", 16))
        self.item_textbox.grid(row=18, column=0, pady=5, rowspan=8)
        self.item_textbox.configure(state="disabled")

        self.unit_label = CTkLabel(self.user_interface_frame, text="Unit (pcs., kgs, boxes, etc.): ")
        self.unit_entry = CTkEntry(self.user_interface_frame, width=275)
        self.unit_label.grid(row=13, column=1, padx=5)
        self.unit_entry.grid(row=14, column=1, padx=5)

        self.article_label = CTkLabel(self.user_interface_frame, text="Article: ")
        self.article_entry = CTkEntry(self.user_interface_frame, width=275)
        self.article_label.grid(row=15, column=1, padx=5)
        self.article_entry.grid(row=16, column=1, padx=5)

        self.item_textbox_label = CTkLabel(self.user_interface_frame, text="Item List")
        self.item_textbox_label.grid(row=17, column=1, pady=5)

        def update_description(choice):
            # amount??
            # amount = quantity * unit price
            try:
                self.quantity = self.items[choice]["quantity"]
                self.unit_price = self.items[choice]["unit_price"]
                self.unit = self.items[choice]["unit"]
                self.amount = self.items[choice]["amount"]
            except Exception as e:
                print(e)
                return
            self.item_textbox.configure(state="normal")
            self.item_textbox.delete("1.0", "end")
            self.item_textbox.insert("end", "Article: " + choice + "\n")
            self.item_textbox.insert("end", "Item Quantity: " + str(self.quantity) + "\n")
            self.item_textbox.insert("end", "Amount: " + str(self.amount) + "\n")
            self.item_textbox.insert("end", "Unit: " + str(self.unit) + "\n")
            self.item_textbox.insert("end", "Unit Price: " + str(self.unit_price) + "\n")
            self.item_textbox.configure(state="disabled")
            return

        def optionmenu_callback(choice):
            update_description(choice)
            return


        def add_item():
            self.quantity = self.quantity_entry.get()
            self.unit_price = self.unit_price_entry.get()
            self.unit = self.unit_entry.get()
            self.article = self.article_entry.get()

            if not self.quantity or not self.unit_price or not self.unit or not self.article:
                mb.showerror("Item Adder: Empty Fields", "Please make sure all fields are filled out.")
                return

            try:
                self.quantity = int(self.quantity)
                if self.quantity <= 0:
                    mb.showerror("Item Adder: Quantity equal to 0", "Please enter a quantity greater than 0.")
                    return
            except:
                mb.showerror("Item Adder: Invalid Quantity", "Please enter a valid quantity greater than 0. Omit all nondigit characters in the input field.")
                return

            try:
                self.unit_price = float(self.unit_price)
                if self.unit_price <= 0:
                    mb.showerror("Item Adder: Unit Price equal to 0", "Please enter a unit price greater than 0.")
                    return
            except:
                mb.showerror("Item Adder: Invalid Unit Price", "Please enter a valid unit price greater than 0. Omit all nondigit characters in the input field.")
                return

            for name in self.items.keys():
                if name == self.article:
                    mb.showerror("Item Adder: Article already exists", "That article already exists in the catalog. Please delete the pre-existing article and create a new one if you wish to make changes.")
                    return
            
            self.amount = self.quantity * self.unit_price
            item = {
                "quantity": self.quantity,
                "unit": self.unit,
                "unit_price": self.unit_price,
                "amount": float(self.amount)
            }
            self.items[self.article] = item
            self.item_selection.configure(values=self.items.keys())
            self.item_selection.set(self.article)
            update_description(self.article)
            return

        def remove_item():
            self.selected = self.item_selection.get()
            del self.items[self.selected]
            self.item_textbox.configure(state="normal")
            self.item_textbox.delete("1.0", "end")
            self.item_textbox.configure(state="disabled")
            if len(self.items) <= 0:
                self.item_selection.set("Add an item")
            else:
                self.item_selection.configure(values=self.items.keys())
                update_description(self.article)
            return

        self.item_selection = CTkOptionMenu(self.user_interface_frame, values=[], fg_color="#343638", button_color="#343638", button_hover_color="#565b5e", command=optionmenu_callback)
        self.item_selection.set("Add an item")
        self.item_selection.grid(row=18, column=1, pady=5)

        self.add_item_button = CTkButton(self.user_interface_frame, text="Add Item", fg_color="#343638", border_color="#565b5e", width=275, command=add_item)
        self.add_item_button.grid(row=19, column=1, pady=5)

        self.remove_item_button = CTkButton(self.user_interface_frame, text="Remove Item", fg_color="#343638", border_color="#565b5e", width=275, command=remove_item)
        self.remove_item_button.grid(row=20, column=1, pady=5)

        self.record_receipt_button = CTkButton(self.user_interface_frame, text="Record Receipt", fg_color="#343638", border_color="#565b5e", width=275, command=self.record_receipt)
        self.record_receipt_button.grid(row=21, column=1, pady=5)
        return
