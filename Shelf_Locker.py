from Global import grabobjs
from Global import ShelfHandle
from tkinter import *
from tkinter import messagebox

import os
import random
import pandas as pd

CurrDir = os.path.dirname(os.path.abspath(__file__))
PreserveDir = os.path.join(CurrDir, '04_Preserve')
ShelfObj = ShelfHandle(os.path.join(PreserveDir, 'Data_Locker'))
Global_Objs = grabobjs(CurrDir)


class MainGUI:
    listbox = None
    obj = None

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
                    mylist.append([item[0], item[1], item[2] + '_' + str(num), item[3], item[5]])
                    pd.DataFrame([item[3]]).to_excel(writer, index=False, header=False,
                                                     sheet_name=item[2] + '_' + str(num))
                    item[4].to_excel(writer, index=False, startrow=1, sheet_name=item[2] + '_' + str(num))

                    num += 1

                df = pd.DataFrame(mylist, columns=['Orig_File_Name', 'File_Creator_Name', 'Tab_Name', 'SQL_Table',
                                                   'Append_Time'])
                df.to_excel(writer, index=False, sheet_name='TAB_Details')

            sys.exit()

        else:
            messagebox.showerror('Selection Error!', 'No shelf date was selected. Please select a valid shelf item')

    def settings(self):
        if self.obj:
            self.obj.close()

        self.obj = SettingsGUI(self.root)
        self.obj.buildgui()

    def cancel(self):
        if self.obj:
            self.obj.close()

        self.root.destroy()


class SettingsGUI:
    entry1 = None
    entry2 = None
    radio1 = None
    radio2 = None

    def __init__(self, root):
        self.dialog = Toplevel(root)
        self.rvar = IntVar()
        self.evar = StringVar()
        self.local_settings = Global_Objs['Local_Settings'].grab_list()

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
        self.entry1.bind('<KeyRelease>', self.checkshelf)

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

        btn = Button(self.dialog, text="Save Settings", width=10, command=self.save_settings)
        btn2 = Button(self.dialog, text="Cancel", width=10, command=self.close)
        btn.pack(in_=bottom_frame, side=LEFT, padx=10)
        btn2.pack(in_=bottom_frame, side=RIGHT, padx=10)
        self.dialog.mainloop()

    def checkshelf(self, event):
        if self.entry1.get() in self.local_settings and self.entry1.get() != "General_Settings_Path":
            myitems = self.local_settings[self.entry1.get()]

            if myitems[0]:
                self.radio1.select()
            else:
                self.radio2.select()

            self.entry2.delete(0, len(self.entry2.get()))
            self.entry2.insert(0, myitems[1])

    def save_settings(self):
        if self.entry1.get() == "General_Settings_Path":
            messagebox.showerror('Table Error', 'Unable to have General_Settings_Path as a table name')
        else:
            if self.rvar.get() == 1:
                myitems = [True, self.entry2.get()]
            else:
                myitems = [False, self.entry2.get()]

            if self.entry1.get() in self.local_settings:
                Global_Objs['Local_Settings'].del_item(self.entry1.get())

            if self.rvar.get() != 1 or self.entry2.get() != '14':
                Global_Objs['Local_Settings'].add_item(self.entry1.get(), myitems)

            self.dialog.destroy()

    def close(self):
        self.dialog.destroy()


if __name__ == '__main__':
    myobj = MainGUI()
    myobj.buildgui()
    myobj.showgui()
