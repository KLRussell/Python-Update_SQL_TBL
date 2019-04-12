from Global import ShelfHandle
from tkinter import *
from tkinter import messagebox

import os
import random
import pandas as pd

CurrDir = os.path.dirname(os.path.abspath(__file__))
PreserveDir = os.path.join(CurrDir, '04_Preserve')
ShelfObj = ShelfHandle(os.path.join(PreserveDir, 'Data_Locker'))


class MainGUI:
    listbox = None

    def __init__(self):
        self.root = Tk()
        self.var = StringVar()

    def buildgui(self):
        self.root.title('Shelf Locker')
        top = Frame(self.root)
        middle = Frame(self.root)
        bottom = Frame(self.root)
        top.pack(side=TOP)
        middle.pack(side=TOP)
        bottom.pack(side=BOTTOM, fill=BOTH, expand=True)

        text = Message(self.root, textvariable=self.var, width=180, justify=CENTER)
        self.var.set('Please choose a date to export Shelved content:')
        text.pack(in_=top)

        self.listbox = Listbox(self.root, selectmode=SINGLE, width=35, yscrollcommand=True)
        self.populatebox()
        self.listbox.pack(in_=middle)

        btn = Button(self.root, text="Get Shelf", width=7, command=self.extract_shelf)
        btn2 = Button(self.root, text="Cancel", width=7, command=self.cancel)
        btn3 = Button(self.root, text="Settings", width=7, command=self.settings)
        btn.pack(in_=bottom, side=LEFT)
        btn2.pack(in_=bottom, side=RIGHT)
        btn3.pack(in_=bottom, side=TOP)

    def showgui(self):
        self.root.mainloop()

    def populatebox(self):
        mykeys = ShelfObj.get_keys()

        for i in mykeys:
            self.listbox.insert("end", i)

    def extract_shelf(self):
        if self.listbox.curselection():
            mylist = []
            export_dir = os.path.join(PreserveDir, 'Data_Locker_Export')
            selection = self.listbox.get(self.listbox.curselection())
            myitems = ShelfObj.grab_item(selection)
            filepath = os.path.join(export_dir, '{0}_{1}.xlsx'.format(
                selection, random.randint(10000000, 100000000)))

            if not os.path.exists(export_dir):
                os.makedirs(export_dir)

            num = 1

            with pd.ExcelWriter(filepath) as writer:
                for item in myitems:
                    print(item)
                    mylist.append([item[0], item[1] + '_' + str(num), item[2], item[4]])
                    item[3].to_excel(writer, sheet_name=item[1] + '_' + str(num))

                    num += 1

                df = pd.DataFrame(mylist, columns=['File_Creator_Name', 'Tab_Name', 'SQL_Table', 'Append_Time'])
                df.to_excel(writer, sheet_name='TAB_Details')

            sys.exit()

        else:
            messagebox.showerror('Selection Error!', 'No shelf date was selected. Please select a valid shelf item')

    @staticmethod
    def settings():
        myobj = SettingsGUI()
        myobj.buildgui()

    @staticmethod
    def cancel():
        sys.exit()


class SettingsGUI:
    entry1 = None
    entry2 = None
    radio1 = None
    radio2 = None

    def __init__(self):
        self.dialog = Tk()
        self.rvar = IntVar()

    def buildgui(self):
        self.dialog.geometry('250x200+500+300')
        self.dialog.title('Update TBL Settings')

        top_frame = Frame(self.dialog)
        middle_frame = Frame(self.dialog)
        middle_frame2 = Frame(self.dialog)
        middle_frame3 = Frame(self.dialog)
        bottom_frame = Frame(self.dialog)
        top_frame.pack(side=TOP)
        middle_frame.pack()
        middle_frame2.pack()
        middle_frame3.pack()
        bottom_frame.pack(side=BOTTOM, fill=BOTH, expand=True)

        header = Message(self.dialog, text='Please input custom settings for individual SQL Server tables:', width=240)
        header.pack(in_=top_frame)

        label1 = Label(self.dialog, text='SQL Server TBL: ', pady=7)
        self.entry1 = Entry(self.dialog)
        label1.pack(in_=middle_frame, side=LEFT)
        self.entry1.pack(in_=middle_frame, side=RIGHT)

        self.radio1 = Radiobutton(self.dialog, text='Autofill Edit_DT On', variable=self.rvar, value=1, pady=5)
        self.radio1.pack(in_=middle_frame2, anchor=W)
        self.radio2 = Radiobutton(self.dialog, text='Autofill Edit_DT Off', variable=self.rvar, value=2, pady=1)
        self.radio2.pack(in_=middle_frame2, anchor=W)
        self.radio1.select()

        label2 = Label(self.dialog, text='Shelf Life Days: ', pady=9)
        self.entry2 = Entry(self.dialog, width=5)
        label2.pack(in_=middle_frame3, side=LEFT)
        self.entry2.pack(in_=middle_frame3, side=RIGHT)
        self.entry2.insert(0, "14")


if __name__ == '__main__':
    myobj = MainGUI()
    myobj.buildgui()
    myobj.showgui()
