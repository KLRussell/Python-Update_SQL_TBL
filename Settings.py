# Global Module import
from tkinter import *
from tkinter import messagebox
from Global import grabobjs
from Global import CryptHandle
from Global import ShelfHandle

import os
import random
import pandas as pd

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
global_objs = grabobjs(main_dir, 'TSQL')
preserve_dir = os.path.join(main_dir, '04_Preserve')
export_dir = os.path.join(preserve_dir, 'Data_Locker_Export')


class SettingsGUI:
    shelf_obj = None
    sql_tbl_txtbox = None
    shelf_file_txtbox = None
    autofill_radio1 = None
    autofill_radio2 = None
    save_button = None

    # Function that is executed upon creation of SettingsGUI class
    def __init__(self):
        self.header_text = 'Welcome to TSQL Table Settings!\nSettings can be changed below.\nPress save when finished'

        self.asql = global_objs['SQL']
        self.local_settings = global_objs['Local_Settings'].grab_list()
        self.main = Tk()

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.sql_tbl = StringVar()
        self.autofill = IntVar()
        self.shelf_life = IntVar()

        # GUI Bind On Destruction event
        self.main.bind('<Destroy>', self.gui_destroy)

    # Function executes when GUI is destroyed
    def gui_destroy(self, event):
        self.asql.close()

    # Static function to fill textbox in GUI
    @staticmethod
    def fill_textbox(setting_list, val, key):
        assert (key and val and setting_list)
        item = global_objs[setting_list].grab_item(key)

        if isinstance(item, CryptHandle):
            val.set(item.decrypt_text())

    # static function to add setting to Local_Settings shelf files
    @staticmethod
    def add_setting(setting_list, val, key, encrypt=True):
        assert (key and setting_list)

        global_objs[setting_list].del_item(key)

        if val:
            global_objs[setting_list].add_item(key=key, val=val, encrypt=encrypt)

    # Function to build GUI for settings
    def build_gui(self, header=None):
        # Change to custom header title if specified
        if header:
            self.header_text = header

        # Set GUI Geometry and GUI Title
        self.main.geometry('444x257+500+190')
        self.main.title('TSQL Table Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=444, height=70)
        shelf_frame = LabelFrame(self.main, text='Shelf Life Settings', width=444, height=140)
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        shelf_frame.pack(fill="both")
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)
        server_txtbox.bind('<KeyRelease>', self.check_network)

        #     Server Database Input Box
        database_label = Label(self.main, text='Database:')
        database_txtbox = Entry(self.main, textvariable=self.database)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=7, padx=15)
        database_label.pack(in_=network_frame, side=RIGHT)
        database_txtbox.bind('<KeyRelease>', self.check_network)

        # Apply Shelf Life Settings to the Shelf_Frame
        #     SQL TBL Input Box
        sql_tbl_label = Label(shelf_frame, text='SQL TBL:')
        self.sql_tbl_txtbox = Entry(shelf_frame, textvariable=self.sql_tbl, width=57)
        sql_tbl_label.grid(row=0, column=0, padx=8, pady=5)
        self.sql_tbl_txtbox.grid(row=0, column=1, columnspan=3, padx=5, pady=5)
        self.sql_tbl_txtbox.bind('<KeyRelease>', self.check_shelf)

        #     Autofill Radio buttons
        self.autofill_radio1 = Radiobutton(shelf_frame, text='Autofill Edit_DT On', variable=self.autofill, value=1)
        self.autofill_radio2 = Radiobutton(shelf_frame, text='Autofill Edit_DT Off', variable=self.autofill, value=2)
        self.autofill_radio1.grid(row=1, column=1, padx=8, pady=5)
        self.autofill_radio2.grid(row=1, column=2, padx=8, pady=5)

        #     Shelf Life Input Box
        shelf_file_label = Label(shelf_frame, text='Shelf Life:')
        self.shelf_file_txtbox = Entry(shelf_frame, textvariable=self.shelf_life, width=57)
        shelf_file_label.grid(row=2, column=0, padx=8, pady=5, sticky=E)
        self.shelf_file_txtbox.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky=W)

        # Apply Buttons to Button_Frame
        #     Save Button
        self.save_button = Button(self.main, text='Save Settings', width=15, command=self.save_settings)
        self.save_button.pack(in_=button_frame, side=LEFT, padx=10, pady=5)

        #     Cancel Button
        extract_button = Button(self.main, text='Cancel', width=15, command=self.cancel)
        extract_button.pack(in_=button_frame, side=RIGHT, padx=10, pady=5)

        #     Extract Shelf Button
        extract_button = Button(self.main, text='Extract Shelf', width=15, command=self.extract_shelf)
        extract_button.pack(in_=button_frame, side=TOP, padx=10, pady=5)

        # Fill Textboxes with settings
        self.fill_gui()

        # Show GUI Dialog
        self.main.mainloop()

    # Function to fill GUI textbox fields
    def fill_gui(self):
        self.fill_textbox('Settings', self.server, 'Server')
        self.fill_textbox('Settings', self.database, 'Database')
        self.autofill.set(1)
        self.shelf_life.set(14)

        if not self.server.get() or not self.database.get() or not self.asql.test_conn('alch'):
            self.sql_tbl_txtbox.configure(state=DISABLED)
            self.save_button.configure(state=DISABLED)
            self.shelf_file_txtbox.configure(state=DISABLED)
            self.autofill_radio1.configure(state=DISABLED)
            self.autofill_radio2.configure(state=DISABLED)
        else:
            self.asql.connect('alch')

    # Function to check network settings if populated
    def check_network(self, event):
        if self.server.get() and self.database.get() and \
                (global_objs['Settings'].grab_item('Server') != self.server.get() or
                 global_objs['Settings'].grab_item('Database') != self.database.get()):
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.asql.test_conn('alch'):
                self.sql_tbl_txtbox.configure(state=NORMAL)
                self.save_button.configure(state=NORMAL)
                self.shelf_file_txtbox.configure(state=NORMAL)
                self.autofill_radio1.configure(state=NORMAL)
                self.autofill_radio2.configure(state=NORMAL)
                self.add_setting('Settings', self.server.get(), 'Server')
                self.add_setting('Settings', self.database.get(), 'Database')
                self.asql.connect('alch')

    # Function to validate whether a SQL table exists in SQL server
    def check_table(self, table):
        table2 = table.split('.')

        if len(table2) == 2:
            myresults = self.asql.query('''
                SELECT
                    1
                FROM information_schema.tables
                WHERE
                    Table_Schema = '{0}'
                        AND
                    Table_Name = '{1}'
            '''.format(table2[0], table2[1]))

            if myresults.empty:
                return False
            else:
                return True
        else:
            return False

    # Function to check shelf depository to see if sql_tbl exists in depository
    #   This function will load settings if sql tbl exists in depository
    def check_shelf(self, event):
        if self.sql_tbl.get() in self.local_settings and self.sql_tbl.get() != 'General_Settings_Path':
            myitems = self.local_settings[self.sql_tbl.get()]

            if myitems[0]:
                self.autofill_radio1.select()
            else:
                self.autofill_radio2.select()

            self.shelf_life.set(myitems[1])

    # Function to connect to SQL connection for this class
    def sql_connect(self):
        if self.asql.test_conn('alch'):
            self.asql.connect('alch')
            return True
        else:
            return False

    # Function to close SQL connection for this class
    def sql_close(self):
        self.asql.close()

    # Function to save settings when the Save Settings button is pressed
    def save_settings(self):
        if self.server.get() and self.database.get():
            if not self.sql_tbl.get():
                messagebox.showerror('SQL TBL Empty Error!', 'No value has been inputed for SQL TBL',
                                     parent=self.main)
            elif not self.shelf_life.get():
                messagebox.showerror('Shelf Life Empty Error!', 'No value has been inputed for Shelf Life',
                                     parent=self.main)
            elif self.shelf_life.get() <= 0:
                messagebox.showerror('Invalid Shelf Life Error!', 'Shelf Life <= 0',
                                     parent=self.main)
            else:
                if not self.check_table(self.sql_tbl.get()):
                    messagebox.showerror('Invalid SQL TBL!',
                                         'SQL TBL does not exist in sql server',
                                         parent=self.main)
                else:
                    if self.autofill.get() == 1:
                        myitems = [True, self.shelf_life.get()]
                    else:
                        myitems = [False, self.shelf_life.get()]

                    if self.autofill.get() != 1 or self.shelf_life.get() != 14:
                        self.add_setting('Local_Settings', myitems, self.sql_tbl.get(), False)
                    else:
                        global_objs['Local_Settings'].del_item(self.sql_tbl.get())

                    self.main.destroy()

    # Function to load extract Shelf GUI
    def extract_shelf(self):
        if self.shelf_obj:
            self.shelf_obj.cancel()

        self.shelf_obj = ExtractShelf(self.main)
        self.shelf_obj.build_gui()

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


class ExtractShelf:
    save_button = None
    list_box = None
    list_sel = 0

    # Function that is executed upon creation of ExtractShelf class
    def __init__(self, root):
        self.shelf_obj = ShelfHandle(os.path.join(preserve_dir, 'Data_Locker'))
        self.main = Toplevel(root)
        self.header_text = 'Welcome to Shelf Date Extraction!\nPlease choose a date below.\nWhen finished press extract'

    # Function to build GUI for Extract Shelf
    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('252x280+500+190')
        self.main.title('Shelf Extractor')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        shelf_frame = LabelFrame(self.main, text='Shelf Locker', width=444, height=140)
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        shelf_frame.pack(fill="both")
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Listbox with scrollbar to the Shelf_Frame
        yscrollbar = Scrollbar(self.main, orient='vertical')
        self.list_box = Listbox(self.main, selectmode=SINGLE, width=35, yscrollcommand=yscrollbar.set)
        yscrollbar.config(command=self.list_box.yview)
        self.list_box.pack(in_=shelf_frame, side=LEFT, padx=5, pady=5)
        yscrollbar.pack(in_=shelf_frame, side=LEFT, fill=Y, pady=5)
        self.list_box.bind("<Down>", self.on_list_down)
        self.list_box.bind("<Up>", self.on_list_up)
        self.list_box.bind('<<ListboxSelect>>', self.on_select)

        # Apply Buttons to Button_Frame
        #     Save Button
        self.save_button = Button(self.main, text='Extract', width=12, command=self.extract_shelf)
        self.save_button.pack(in_=button_frame, side=LEFT, padx=10, pady=5)

        #     Cancel Button
        extract_button = Button(self.main, text='Cancel', width=12, command=self.cancel)
        extract_button.pack(in_=button_frame, side=RIGHT, padx=10, pady=5)

        self.load_gui()

    def load_gui(self):
        for key in self.shelf_obj.get_keys():
            self.list_box.insert('end', key)

        if self.list_box.size() > 0:
            self.list_box.select_set(0)
        else:
            self.save_button.configure(state=DISABLED)

    # Function adjusts selection of item when user clicks item
    def on_select(self, event):
        if self.list_box and self.list_box.curselection() and -1 < self.list_sel < self.list_box.size() - 1:
            self.list_sel = self.list_box.curselection()[0]

    def on_list_down(self, event):
        if self.list_sel < self.list_box.size() - 1:
            self.list_box.select_clear(self.list_sel)
            self.list_sel += 1
            self.list_box.select_set(self.list_sel)

    def on_list_up(self, event):
        if self.list_sel > 0:
            self.list_box.select_clear(self.list_sel)
            self.list_sel -= 1
            self.list_box.select_set(self.list_sel)

    def extract_shelf(self):
        if self.list_box.size() > 0:
            if self.list_box.curselection():
                mylist = []
                selection = self.list_box.get(self.list_box.curselection())
                myitems = self.shelf_obj.grab_item(selection)
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

                    df = pd.DataFrame(mylist, columns=['Orig_File_Name', 'File_Authors', 'Tab_Name', 'SQL_Table',
                                                       'Append_Time'])
                    df.to_excel(writer, index=False, sheet_name='TAB_Details')

                self.main.destroy()
            else:
                messagebox.showerror('Selection Error!', 'No shelf date was selected. Please select a valid shelf item')

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


# Main loop routine to create GUI Settings
if __name__ == '__main__':
    obj = SettingsGUI()

    try:
        obj.build_gui()
    finally:
        obj.sql_close()
