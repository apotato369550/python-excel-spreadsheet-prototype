import os
from customtkinter import *
import customtkinter as ctk
from .POUI import POUI
from .ARARUI import ARARUI

class CustomUI(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")  
        self.max_width = 500
        self.max_height = 225

        self.title("InvoiceMate Version 1.0")
        self.geometry(str(self.max_width) + "x" + str(self.max_height))
        self.resizable(False, False)
 
        self.make_user_interface()
        self.mainloop()
        pass

    def make_user_interface(self):
        self.user_interface_frame = CTkFrame(self, width=self.max_width, height=self.max_height)
        self.user_interface_frame.pack()

        self.title_label = CTkLabel(self.user_interface_frame, text="InvoiceMate Version 1.0", font=("Courier", 30, "bold"))
        self.title_label.grid(row=0, column=0, padx=45)
        self.title_label = CTkLabel(self.user_interface_frame, text="Developed by: Jay and CJ", font=("Courier", 10))
        self.title_label.grid(row=1, column=0)

        # add commands here
        self.open_receipt_manager_button = CTkButton(self.user_interface_frame, text="Open Receipt (PO) Manager", font=("Lucida Console", 15), command=lambda: POUI(), fg_color="#343638", border_color="#565b5e")
        self.open_receipt_manager_button.grid(row=2, column=0, pady=5)

        self.open_receivables_manager_button = CTkButton(self.user_interface_frame, text="Open Receivables (ARAR) Manager", font=("Lucida Console", 15), command=lambda: ARARUI(), fg_color="#343638", border_color="#565b5e")
        self.open_receivables_manager_button.grid(row=3, column=0, pady=5)
        
        self.open_readme_button = CTkButton(self.user_interface_frame, text="View Readme/Help Guide", font=("Lucida Console", 15), command=self.open_readme, fg_color="#343638", border_color="#565b5e")
        self.open_readme_button.grid(row=4, column=0, pady=5)

        self.exit_button = CTkButton(self.user_interface_frame, text="Exit Program", font=("Lucida Console", 15), command=lambda: exit(), fg_color="#343638", border_color="#565b5e")
        self.exit_button.grid(row=5, column=0, pady=5)
        return

    def open_readme(self):
        return
