import pygetwindow as gw
import pyautogui
import getpass
from pywinauto import Application
from datetime import datetime
from tkinter import messagebox

import time

username = getpass.getuser()
today = datetime.today().strftime('%d.%m.%Y')


def first_transaction_execute(session, sales_job, vendor_number, order_price, asc_responsible):
    with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
        log_file.write(f'\n[{datetime.today()}] - {"*"*5} ORDER: {sales_job} - {asc_responsible} {"*"*5}\n')

        session.findById("wnd[0]/tbar[0]/okcd").text = "iw32"
        log_file.write(f'[{datetime.today()}] - Going to IW32 transaction (Initial Screen)\n')
        print("Going to IW32 transaction (Initial Screen)")

        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print("Pressed Enter")

        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = sales_job
        log_file.write(f'[{datetime.today()}] - Order -> {sales_job}\n')
        print(f"Order -> {sales_job}")

        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print("Pressed Enter")

        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").select()
        log_file.write(f'[{datetime.today()}] - Going to "Operations" tab\n')
        print('Going to "Operations" tab')

        try:
            control_key = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-STEUS[4,0]").text
        except:
            control_key = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-STEUS[4,0]").text

        if control_key != "SM02":
            try:
                session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-STEUS[4,0]").text = "SM02"
            except:
                session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-STEUS[4,0]").text = "SM02"
                log_file.write(f'[{datetime.today()}] - Changed "Control Key" (5) column to be "SM02"\n')
                print('Changed "Control Key" (5) column to be "SM02"')

            session.findById("wnd[0]").sendVKey(0)
            log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
            print("Pressed Enter")
        else:
            log_file.write(f'[{datetime.today()}] - "Control Key" (5) column is already "SM02", leaving as it is...\n')

        print('"Control Key" (5) column is already "SM02", leaving as it is...')

        try:
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_VGD2").press()
        except:
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_VGD2").press()

        log_file.write(f'[{datetime.today()}] - Clicked "External"\n')
        print('Clicked "External"')
        try:
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1200/subSUB_AVO_TABSTRIP:SAPLCOIH:1205/tabsTS_1205/tabpVGD2/ssubSUB_AVO_DETAIL:SAPLCOIH:1220/txtAFVGD-PREIS").text = order_price
            log_file.write(f'[{datetime.today()}] - Price -> {order_price}\n')
            print(f'Price -> {order_price}')
        except:
            return "Blocked"

        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1200/subSUB_AVO_TABSTRIP:SAPLCOIH:1205/tabsTS_1205/tabpVGD2/ssubSUB_AVO_DETAIL:SAPLCOIH:1220/ctxtAFVGD-LIFNR").text = vendor_number
        log_file.write(f'[{datetime.today()}] - Vendor -> {vendor_number}\n')
        print('Vendor -> {vendor_number}')

        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print('Pressed Enter')

        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        log_file.write(f'[{datetime.today()}] - Saving transaction\n')
        print('Saving Transaction')

        time.sleep(2)
        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(12)
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").select()
        log_file.write(f'[{datetime.today()}] - Going to "Operations" tab\n')
        print(13)
        try:
            requisition_number = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-BANFN[7,0]").text
        except:
            requisition_number = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-BANFN[7,0]").text

        log_file.write(f'[{datetime.today()}] - "Requisition" (8) column is {requisition_number}\n')
        print(14)
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        log_file.write(f'[{datetime.today()}] - Saving transaction\n')
        print(15)
        print(requisition_number)
        return requisition_number


def second_transaction_execute(session, requisition_number):
    with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nme21n"
        log_file.write(f'[{datetime.today()}] - Going to ME21N transaction (Initial Screen)\n')
        print(1)
        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(2)
        try:
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[4,0]").text = requisition_number
        except:
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[4,0]").text = requisition_number
        log_file.write(f'[{datetime.today()}] - Entered in column "Purchase Req." (5) -> {requisition_number}\n')
        print(3)
        session.findById("wnd[0]").sendVKey(0)
        try:
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[7,0]").text = 1
        except:
            try:
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[7,0]").text = 1
            except:
                try:
                    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[7,0]").text = 1
                except:
                    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[7,0]").text = 1

        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(4)

        try:
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7").select()
        except:
            try:
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7").select()
            except:
                try:
                    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7").select()
                except:
                    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7").select()

        log_file.write(f'[{datetime.today()}] - Going to "Delivery" tab\n')

        print(5)

        try:
            tax_code = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ").text
        except:
            try:
                tax_code = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ").text
            except:
                try:
                    tax_code = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ").text
                except:
                    tax_code = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ").text

        print(5.5)

        if tax_code != "MH":
            try:
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ").text = "MH"
            except:
                try:
                    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ").text = "MH"
                except:
                    try:
                        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ").text = "MH"
                    except:
                        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ").text = "MH"

            log_file.write(f'[{datetime.today()}] - Changed tax code from {tax_code} to MH\n')
        else:
            log_file.write(f'[{datetime.today()}] - Tax code is already MH, leaving as it is...\n')
        print(6)
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        log_file.write(f'[{datetime.today()}] - Saving transaction\n')
        print(7)
        order_number_text = session.findById("wnd[0]/sbar/pane[0]").text
        if order_number_text == '' or order_number_text is None:
            log_file.write(f'[{datetime.today()}] - ERROR: No order number was generated!\n')
            return
        else:
            log_file.write(f'[{datetime.today()}] - {order_number_text}\n')
        print(8)
        try: return int(order_number_text.split(' ')[-1])
        except: print(order_number_text)

"""
def third_transaction_execute(session, order_number):
    with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nme22n"
        log_file.write(f'[{datetime.today()}] - Going to ME22N transaction\n')
        print(1)
        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(2)
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = order_number
        session.findById("wnd[0]").sendVKey(0)
        print(3)
        session.findById("wnd[0]/tbar[1]/btn[20]").press()
        log_file.write(f'[{datetime.today()}] - Clicked "Print Preview"\n')
        print(4)
        session.findById("wnd[0]/tbar[0]/okcd").text = "pdf!"
        log_file.write(f'[{datetime.today()}] - Entered "pdf!" in the search bar\n')
        print(5)
        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(6)
        time.sleep(1.5)

        while not gw.getWindowsWithTitle('PDF Preview')[0]:
            pass

        pdf_preview_window = gw.getWindowsWithTitle('PDF Preview')[0]
        log_file.write(f'[{datetime.today()}] - Found the "PDF Preview" window\n')
        print(7)

        title_bar_x = pdf_preview_window.left + pdf_preview_window.width // 2
        title_bar_y = pdf_preview_window.top + 5

        pyautogui.click(x=title_bar_x, y=title_bar_y + 200)
        log_file.write(f'[{datetime.today()}] - Clicked on "PDF Preview" window\n')
        print(8)
        pyautogui.hotkey('ctrl', 'shift', 's')
        log_file.write(f'[{datetime.today()}] - Pressed "Save" button\n')
        print(9)

        while True:
            try:
                app = Application().connect(class_name='#32770', title='Save As')
                break
            except:
                pyautogui.click(x=title_bar_x, y=title_bar_y + 200)
                pyautogui.hotkey('ctrl', 'shift', 's')

        save_as_dialogs = app.windows(class_name='#32770', title='Save As', visible_only=True)

        while True:
            try:
                save_as_dialog = save_as_dialogs[0]
                break
            except IndexError:
                save_as_dialogs = app.windows(class_name='#32770', title='Save As', visible_only=True)

        log_file.write(f'[{datetime.today()}] - Found the "Save as" filedialog window\n')
        print(10)

        time.sleep(1.5)
        pyautogui.write(rf"{order_number}.pdf")
        log_file.write(f'[{datetime.today()}] - As the name of the file entered order number -> {order_number}.pdf\n')
        print(11)
        for element in save_as_dialog.descendants():
            if 'Address:' in element.window_text():
                address_edit = element
                print("Address found")
            if '&Save' in element.window_text():
                save_button = element
                print('Save found')
        address_edit.click()
        pyautogui.write(rf"C:\\Users\\{username}\\Desktop\\PO's UK\\Hydraulic Crane Services")
        log_file.write(f'[{datetime.today()}] - As an address of the file entered G:\\Shared drives\\Hiab Data Governance\\CEX-PLGDN1 Documentation Management\\2. Sales Coordination and Aftersales Support (CSC)\\HANDOVERS\\Automated Handover PO\'s\n')
        print(12)
        pyautogui.press("Enter")
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(13)
        time.sleep(1.5)
        save_button.click()
        pyautogui.press("Enter")
        log_file.write(f'[{datetime.today()}] - Clicked "Save"\n')
        print(14)
        session.findById("wnd[1]").close()
        log_file.write(f'[{datetime.today()}] - Closed "PDF Preview" window\n')
        print(15)
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        log_file.write(f'[{datetime.today()}] - Going back (ME22N)\n')
        print(16)
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        log_file.write(f'[{datetime.today()}] - Going back (SAP User Menu)\n')
        print(17)
"""

def fourth_transaction_execute(session, order_number):
    with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nmigo"
        log_file.write(f'[{datetime.today()}] - Going to MIGO transaction\n')
        print(1)
        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(2)
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER").text = order_number
        log_file.write(f'[{datetime.today()}] - Order Number -> {order_number}\n')
        print(3)
        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(4)
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = True
        log_file.write(f'[{datetime.today()}] - Ticked the checkbox Item OK\n')
        print(5)
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").setFocus()
        print(6)
        session.findById("wnd[0]/tbar[1]/btn[7]").press()
        log_file.write(f'[{datetime.today()}] - Clicked "Check"\n')
        print(7)
        session.findById("wnd[0]/tbar[1]/btn[23]").press()
        log_file.write(f'[{datetime.today()}] - Clicked "Post"\n')
        print(8)
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        log_file.write(f'[{datetime.today()}] - Going back\n')
        print(9)


def fifth_transaction_execute(session, sales_job):
    with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
        session.findById("wnd[0]/tbar[0]/okcd").text = "iw32"
        log_file.write(f'[{datetime.today()}] - Going to IW32 transaction (Initial Screen)\n')
        print(1)
        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Order -> {sales_job}\n')
        session.findById("wnd[0]").sendVKey(0)
        log_file.write(f'[{datetime.today()}] - Pressed Enter\n')
        print(2)
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpKOAU").select()
        log_file.write(f'[{datetime.today()}] - Going to "Costs" tab\n')
        print(3)
        try:
            session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press()
        except:
            try:
                session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001").press()
            except:
                session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/subSUB_KOPF:SAPLCOIH:1108/btn%#AUTOTEXT001").press()
        log_file.write(f'[{datetime.today()}] - Clicked on the "pen with the green tick" icon\n')
        print(4)
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 1
        print(5)
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 2
        print(6)
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E").verticalScrollbar.position = 3
        log_file.write(f'[{datetime.today()}] - Scrolled down\n')
        print(7)
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,3]").selected = True
        log_file.write(f'[{datetime.today()}] - "80 WDON Work Done" status selected\n')
        print(8)
        session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,3]").setFocus()
        print(9)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        log_file.write(f'[{datetime.today()}] - Clicked on the finish flag icon\n')
        print(10)
        session.findById("wnd[0]/tbar[1]/btn[36]").press()
        log_file.write(f'[{datetime.today()}] - Confirmed clicking green tick icon\n')
        print(11)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        print(12)
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        log_file.write(f'[{datetime.today()}] - Going back\n')
        log_file.write(f"{'-'*20}\n")


def operation_execute(session, sales_job, vendor_number, order_price, asc_responsible):

    try:
        requisition_number = first_transaction_execute(session, sales_job, vendor_number, order_price, asc_responsible)
        if requisition_number == "Blocked":
            with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
                log_file.write(f'[{datetime.today()}] - ERROR in transaction IW32 (first): The order has been already processed, proceeding to the next one...\n\n')
            messagebox.showerror("ERROR in transaction IW32 (first)", "Error: The order has been already processed, proceeding to the next one...")
            print("Blocked")
            return "Blocked"
    except Exception as e:
        with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
            log_file.write(f'[{datetime.today()}] - ERROR in transaction IW32 (first): {e}\n\n')
            messagebox.showerror('The error in IW32 (first)', f"Error: {e}")
            return


    try:
        order_number = second_transaction_execute(session, requisition_number)
    except Exception as e:
        with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
            log_file.write(f'[{datetime.today()}] - ERROR in transaction ME21N: {e}\n')
            messagebox.showerror('The error in ME21N', f"Error: {e}")
            return

    print(order_number)

    """
    try:
        third_transaction_execute(session, order_number)
    except Exception as e:
        with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
            log_file.write(f'[{datetime.today()}] - ERROR in transaction ME22N: {e}\n')
            messagebox.showerror('The error in ME22N', f"Error: {e}")
            return
    """

    try:
        fourth_transaction_execute(session, order_number)
    except Exception as e:
        with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
            log_file.write(f'[{datetime.today()}] - ERROR in transaction MIGO: {e}\n')
            messagebox.showerror('The error in MIGO', f"Error: {e}")
            return

    try:
        fifth_transaction_execute(session, sales_job)
    except Exception as e:
        with open(f"static/logs/{today}_{username}.txt", "a") as log_file:
            log_file.write(f'[{datetime.today()}] - ERROR in transaction IW32 (second): {e}\n')
            messagebox.showerror('The error in IW32 (second)', f"Error: {e}")
            return

    return order_number
