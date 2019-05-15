# Global Module import
from tkinter import *
from tkinter import messagebox
from Global import grabobjs
from Global import CryptHandle
import os

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
global_objs = grabobjs(main_dir)


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
    def add_setting(setting_list, val, key):
        assert (key and val and setting_list)

        global_objs[setting_list].del_item(key)
        global_objs[setting_list].add_item(key=key, val=val, encrypt=True)

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

        if not self.server.get() or not self.database.get():
            self.sql_tbl_txtbox.configure(state=DISABLED)
            self.save_button.configure(state=DISABLED)

        self.shelf_file_txtbox.configure(state=DISABLED)
        self.autofill_radio1.configure(state=DISABLED)
        self.autofill_radio2.configure(state=DISABLED)

    # Function to check network settings if populated
    def check_network(self):
        if self.server.get() and self.database.get() and \
                (global_objs['Settings'].grab_item('Server') != self.server.get() or
                 global_objs['Settings'].grab_item('Database') != self.database.get()):
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.asql.test_conn('alch'):
                self.sql_tbl_txtbox.configure(state=NORMAL)
                self.save_button.configure(state=NORMAL)
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

    def check_shelf(self):
        if self.sql_tbl.get() in self.local_settings and self.sql_tbl.get() != 'General_Settings_Path':
            self.shelf_file_txtbox.configure(state=NORMAL)
            self.autofill_radio1.configure(state=NORMAL)
            self.autofill_radio2.configure(state=NORMAL)
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

                    self.add_setting('Local_Settings', myitems, self.sql_tbl.get())

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
    # Function that is executed upon creation of ExtractShelf class
    def __init__(self, root):
        self.main = Toplevel(root)
        self.header_text = 'Welcome to Shelf Date Extraction!\nPlease choose a date below.\nWhen finished press extract'

    # Function to build GUI for Extract Shelf
    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('444x257+500+190')
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

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


# Main loop routine to create GUI Settings
if __name__ == '__main__':
    obj = SettingsGUI()
    obj.build_gui()

