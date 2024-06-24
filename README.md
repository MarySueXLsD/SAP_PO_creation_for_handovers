# SAP_PO_creation_for_handovers
RPA for purchase order creation in SAP to process handover in the system with robot

**main.py**\n
This file is used to render tkinter UI desktop application as well as to connect to SAP and start the process.

**sheet.py**
The file is used to connect to the specific google sheet file, filter out the data and start processing some rows.

**handover_processing.py**
The main SAP function, which goes through transactions IW32, ME21N, ME22N (currently commented out) and MIGO.

**generate.py**
This file is used one time before delivering the application to the end-user in order to generate encrypted username and password to the config.ini file that **main.py** later will read and decrypt.
