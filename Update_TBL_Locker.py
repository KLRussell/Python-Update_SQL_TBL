from Global import ShelfHandle
from tkinter import *
from tkinter import messagebox

import os
import random
import pandas as pd

CurrDir = os.path.dirname(os.path.abspath(__file__))
PreserveDir = os.path.join(CurrDir, '04_Preserve')
ShelfObj = ShelfHandle(os.path.join(PreserveDir, 'Data_Locker'))


def populatebox():
    mykeys = ShelfObj.get_keys()

    for i in mykeys:
        listbox.insert("end", i)


def button1():
    if listbox.curselection():
        mylist = []
        export_dir = os.path.join(PreserveDir, 'Data_Locker_Export')
        selection = listbox.get(listbox.curselection())
        myitems = ShelfObj.grab_item(selection)
        filepath = os.path.join(export_dir, '{0}_Update_{1}.xlsx'.format(
            selection, random.randint(10000000, 100000000)))

        if not os.path.exists(export_dir):
            os.makedirs(export_dir)

        for item in myitems:
            with pd.ExcelWriter(filepath) as writer:
                mylist.append([item[0], item[1], item[3]])
                item[2].to_excel(writer, sheet_name=item[1])

                df = pd.DataFrame(mylist, columns=['File_Creator_Name', 'Tab_Name', 'Append_Time'])
                df.to_excel(writer, sheet_name='Append_Details')

        sys.exit()

    else:
        messagebox.showerror('Selection Error!', 'No shelf date was selected. Please select a valid shelf item')


def cancel():
    sys.exit()


if __name__ == '__main__':
    root = Tk()
    var = StringVar()

    root.title('Shelf Locker')
    top = Frame(root)
    middle = Frame(root)
    bottom = Frame(root)
    top.pack(side=TOP)
    middle.pack(side=TOP)
    bottom.pack(side=BOTTOM, fill=BOTH, expand=True)

    text = Message(root, textvariable=var, width=180, justify=CENTER)
    var.set('Please choose a date to export Shelved content:')
    text.pack(in_=top)

    listbox = Listbox(root, selectmode=SINGLE, width=35, yscrollcommand=True)
    populatebox()
    listbox.pack(in_=middle)

    btn = Button(root, text="Get Shelf", width=10, command=button1)
    btn2 = Button(root, text="Cancel", width=10, command=cancel)
    btn.pack(in_=bottom, side=LEFT)
    btn2.pack(in_=bottom, side=RIGHT)
    root.mainloop()
