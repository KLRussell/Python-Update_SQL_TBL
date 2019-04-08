from Global import ShelfHandle
import os
from tkinter import *

CurrDir = os.path.dirname(os.path.abspath(__file__))
PreserveDir = os.path.join(CurrDir, '04_Preserve')
ShelfObj = ShelfHandle(os.path.join(PreserveDir, 'Data_Locker'))


def populatebox(listbox):
    mykeys = ShelfObj.get_keys()

    for i in mykeys:
        listbox.insert("end", i)


def button1():
    if listbox.curselection():
        myitems = ShelfObj.grab_item(listbox.get(listbox.curselection()))

        for item in myitems:
            print(item)


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
    populatebox(listbox)
    listbox.pack(in_=middle)

    btn = Button(root, text="Get Shelf", width=10, command=button1)
    btn2 = Button(root, text="Cancel", width=10, command=cancel)
    btn.pack(in_=bottom, side=LEFT)
    btn2.pack(in_=bottom, side=RIGHT)
    root.mainloop()
