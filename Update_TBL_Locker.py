from Global import ShelfHandle
import os
from tkinter import *
import tkinter

CurrDir = os.path.dirname(os.path.abspath(__file__))
PreserveDir = os.path.join(CurrDir, '04_Preserve')
listOfCompanies = [[1, ''], [2, '-'], [3, '@ASK TRAINING PTE. LTD.'], [4, 'AAIS'], [5, 'Ademco'], [6, 'Anacle']]


def populatebox(Lb1):
    for i in listOfCompanies:
        Lb1.insert("end", i)


def button1():
    print('hi')


def cancel():
    print('exit')


if __name__ == '__main__':
    root = Tk()
    var = StringVar()

    top = Frame(root)
    middle = Frame(root)
    bottom = Frame(root)
    top.pack(side=TOP)
    middle.pack(side=TOP)
    bottom.pack(side=BOTTOM, fill=BOTH, expand=True)

    text = Message(root, textvariable=var)
    var.set('Please choose a date to export Shelved content:')
    text.pack(in_=top)

    Lb1 = Listbox(root, selectmode=SINGLE, yscrollcommand=True)
    populatebox(Lb1)
    Lb1.pack(in_=middle)

    btn = Button(root, text="Get Shelf", command=button1)
    btn2 = Button(root, text="Cancel", command=cancel)
    btn.pack(in_=bottom, side=LEFT)
    btn2.pack(in_=bottom, side=RIGHT)
    root.mainloop()
