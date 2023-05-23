import os
from customtkinter import *
import customtkinter as ctk

class ReadmeUI(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")  
        self.max_width = 400
        self.max_height = 600

        self.title("InvoiceMate Version 1.0: User Guide")
        self.geometry(str(self.max_width) + "x" + str(self.max_height))
        self.resizable(False, False)

        self.make_user_interface()
        return

    def display_readme(self):
        with open("readme.txt", "r") as file:
            content = file.read()
            self.readme_textbox.configure(state="normal")
            self.readme_textbox.insert("0.0", content)
            file.close()
        return

   
    def make_user_interface(self):
        self.user_interface_frame = CTkFrame(self, width=self.max_width, height=self.max_height)
        self.user_interface_frame.pack()

        self.readme_textbox = CTkTextbox(self.user_interface_frame, width=self.max_width - 20, height=self.max_height - 20, activate_scrollbars=False)
        self.readme_textbox.grid(row=0, column=0, pady=5, sticky="nsew")

        self.readme_scrollbar = CTkScrollbar(self.user_interface_frame, command=self.readme_textbox.yview)
        self.readme_scrollbar.grid(row=0, column=1, sticky="ns")

        self.readme_textbox.configure(yscrollcommand=self.readme_scrollbar.set)

        self.display_readme()
        return