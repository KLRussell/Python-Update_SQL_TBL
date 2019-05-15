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
    # Function that is executed upon creation of SettingsGUI class
    def __init__(self):
        self.header_text = 'Welcome to TSQL Table Settings!\nSettings can be changed below.\nPress save when finished'

        self.asql = global_objs['SQL']
        self.main = Tk()

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.sql_tbl = StringVar()
        self.autofill = IntVar()
        self.shelf_life = IntVar()

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
        self.main.geometry('444x260+500+190')
        self.main.title('TSQL Table Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=444, height=70)
        shelf_frame = LabelFrame(self.main, text='Shelf Life Settings', width=444, height=140)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        shelf_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)

        #     Server Database Input Box
        database_label = Label(self.main, text='Database:')
        database_txtbox = Entry(self.main, textvariable=self.database)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=7, padx=15)
        database_label.pack(in_=network_frame, side=RIGHT)

        # Apply Shelf Life Settings to the Shelf_Frame
        #     SQL TBL Input Box
        sql_tbl_label = Label(shelf_frame, text='SQL TBL:')
        sql_tbl_txtbox = Entry(shelf_frame, textvariable=self.sql_tbl, width=57)
        sql_tbl_label.grid(row=0, column=0, padx=8, pady=5)
        sql_tbl_txtbox.grid(row=0, column=1, columnspan=3, padx=5, pady=5)

        #     Autofill Radio buttons
        autofill_radio1 = Radiobutton(shelf_frame, text='Autofill Edit_DT On', variable=self.autofill, value=1)
        autofill_radio2 = Radiobutton(shelf_frame, text='Autofill Edit_DT Off', variable=self.autofill, value=2)
        autofill_radio1.grid(row=1, column=1, padx=8, pady=5)
        autofill_radio2.grid(row=1, column=2, padx=8, pady=5)

        #     Shelf Life Input Box
        sql_tbl_label = Label(shelf_frame, text='Shelf Life:')
        sql_tbl_txtbox = Entry(shelf_frame, textvariable=self.shelf_life, width=57)
        sql_tbl_label.grid(row=2, column=0, padx=8, pady=5, sticky=E)
        sql_tbl_txtbox.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky=W)

        # Fill Textboxes with settings
        self.fill_gui()

        # Show GUI Dialog
        self.main.mainloop()

    # Function to fill GUI textbox fields
    def fill_gui(self):
        self.fill_textbox('Settings', self.server, 'Server')
        self.fill_textbox('Settings', self.database, 'Database')

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
        if not self.server.get():
            messagebox.showerror('Server Empty Error!', 'No value has been inputed for Server',
                                 parent=self.main)
        elif not self.database.get():
            messagebox.showerror('Database Empty Error!', 'No value has been inputed for Database',
                                 parent=self.main)
        else:
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.asql.test_conn('alch'):
                self.add_setting('Settings', self.server.get(), 'Server')
                self.add_setting('Settings', self.database.get(), 'Database')

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


# Main loop routine to create GUI Settings
if __name__ == '__main__':
    obj = SettingsGUI()
    obj.build_gui()
