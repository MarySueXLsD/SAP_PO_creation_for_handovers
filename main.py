import base64
import configparser
import subprocess
import sys
import getpass
import win32com.client
import pywintypes
from tkinter import *
from tkinter import messagebox, ttk
from datetime import datetime
from cryptography.fernet import Fernet
from sheet import google_sheet_data


class SapGui:
    def sapOpen(self, connectionName):
        try:
            self.connection = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0)
            popup_error('You have current SAP window opened')
            return
        except pywintypes.com_error:
            pass

        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        try:
            subprocess.Popen(self.path)
        except:
            popup_error(r'No SAP installed on the path "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"')
            return

        while True:
            try:
                self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
                break
            except pywintypes.com_error:
                continue

        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return
        self.application = self.SapGuiAuto.GetScriptingEngine
        self.connection = self.application.OpenConnection(connectionName, True)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()
        self.sapLogin(connectionName)

    def sapLogin(self, connectionName):
        if connectionName == "FLP ONE ECC regression test":
            try:
                self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "011"
                self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = cipher_suite.decrypt(base64.urlsafe_b64decode(config.get('SAP_login', 'username'))).decode('utf-8')
                self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = cipher_suite.decrypt(base64.urlsafe_b64decode(config.get('SAP_login', 'password'))).decode('utf-8')
                self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById("wnd[0]").sendVKey(0)
                try: google_sheet_data(self.session)
                except Exception as e: messagebox.showerror('The error', f"Error: {e}")
                sapObj.sapClose()
                log_file.write(f'\n[{datetime.today()}] - SAP connection closed')
            except:
                print(sys.exc_info()[0])
        elif connectionName == "ONE ECC production (FIP)":
            try:
                self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "011"
                self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = cipher_suite.decrypt(base64.urlsafe_b64decode(config.get('SAP_login', 'username'))).decode('utf-8')
                self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = cipher_suite.decrypt(base64.urlsafe_b64decode(config.get('SAP_login', 'password'))).decode('utf-8')
                self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById("wnd[0]").sendVKey(0)
                try: google_sheet_data(self.session)
                except Exception as e: messagebox.showerror('The error', f"Error: {e}")
                sapObj.sapClose()
                log_file.write(f'\n[{datetime.today()}] - SAP connection closed')
            except:
                print(sys.exc_info()[0])

    def sapClose(self):
        try:
            self.connection.CloseSession('ses[0]')
        except AttributeError as e:
            pass


class tk_commands:
    def main_on_resize(self, event):
        global width, height
        if event.widget == app and (event.widget.winfo_height() != height or event.widget.winfo_width() != width):
            new_geometry = f"{event.width}x{event.height}"
            print(f"Window resized to: {new_geometry}")
            config['main_window_size']['width'] = str(event.width)
            config['main_window_size']['height'] = str(event.height)
            with open('static/config.ini', 'w') as configfile:
                config.write(configfile)
            height = event.widget.winfo_height()
            width = event.widget.winfo_width()


if __name__ == "__main__":
    username = getpass.getuser()
    today = datetime.today().strftime('%d.%m.%Y')

    #fernet_key = b'4iSGC9bqiQ7eM9--sm8WcUhX8hns3hrbRJA31SOSEcU='
    fernet_key = b'g3FpQYUxjqu0MNq0AaTv8gKdL3Nv8zKSC5AOsvYkI2Y='
    cipher_suite = Fernet(fernet_key)

    with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
        log_file.write('\n' + '*'*15)
        log_file.write(f'\n[{datetime.today()}] - Application launched\n')

    config = configparser.ConfigParser()
    config.read('static/config.ini')
    app = Tk()
    logo = PhotoImage(file='static/hiab.png')
    app.title('PO Handovers Automation')
    app.iconphoto(False, logo)

    if not config.has_section('main_window_size'):
        config.add_section('main_window_size')
        config['main_window_size']['width'] = '700'
        config['main_window_size']['height'] = '290'
    width = config['main_window_size']['width']
    height = config['main_window_size']['height']

    app.geometry(f'{width}x{height}')

    window = None

    sapObj = SapGui()
    tk_com = tk_commands()


    def popup_error(error_text):
        with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
            log_file.write(f'\n[{datetime.today()}] - ERROR: {error_text}\n')
        messagebox.showerror('Connection Error', 'Error: ' + error_text)

    def on_closing():
        with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
            log_file.write(f'\n[{datetime.today()}] - Application closed\n')
            app.destroy()

    title_label = Label(app, text="PO Handovers Automation", font=('bold', 14), pady=5)
    title_label.grid(row=0, column=0, columnspan=4)

    SAP_frame = Frame(app, highlightbackground="Black", highlightthickness=2)
    SAP_frame.grid(row=1, column=0, columnspan=4, padx=15)

    connection_label = Label(SAP_frame, text="PO: ", font=('bold', 12), padx=20, pady=10)
    connection_label.grid(row=3, column=0)

    connect_btn_first = Button(SAP_frame, text="FLP ONE ECC regression test", command=lambda: sapObj.sapOpen("FLP ONE ECC regression test")) #state=DISABLED)
    connect_btn_first.grid(row=3, column=1)
    connect_btn_second = Button(SAP_frame, text="ONE ECC production (FIP)", command=lambda: sapObj.sapOpen("ONE ECC production (FIP)"))
    connect_btn_second.grid(row=3, column=2, padx=10)

    separator_SAP2 = ttk.Separator(SAP_frame, orient="horizontal")
    separator_SAP2.grid(row=4, column=0, columnspan=6, sticky="ew")

    app.bind("<Configure>", tk_com.main_on_resize)
    app.protocol("WM_DELETE_WINDOW", on_closing)

    mainloop()
