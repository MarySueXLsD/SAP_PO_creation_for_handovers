# SAP automation: PO creation for handovers
RPA for purchase order creation in SAP to process handover in the SAP system with robot. This repository may be the foundation of any of your SAP

**main.py**

This file serves a dual purpose: rendering the Tkinter UI desktop application and establishing a connection to SAP to initiate processes. The config.ini file enables the application to retain its window size settings after closure.

**sheet.py**

The file is used to connect to the specific google sheet file, filter out the data and start processing some rows.

**handover_processing.py**

The main SAP function, which goes through transactions IW32, ME21N, ME22N (currently commented out) and MIGO.

**generate.py**

This file is used one time before delivering the application to the end-user in order to generate encrypted username and password to the config.ini file that **main.py** later will read and decrypt.

![image](https://github.com/MarySueXLsD/SAP_PO_creation_for_handovers/assets/95324605/808e9365-7a47-4f79-9325-a140c706b3e7)

**static/client_secret_test_googleusercontent.com.json**

This file is generated via IAM in google console (https://console.cloud.google.com/) in order to read and write data in the Google Sheet file.

**static/config.ini**

This file is used for storing the encrypted SAP credentials as well as the tkinter window size.

**static/logs/...txt**

Those files are used to store the logs of each action happened after opening the application.
