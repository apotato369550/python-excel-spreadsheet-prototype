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
        return


if __name__ == "__main__":
    arar_maker = ARARMaker()