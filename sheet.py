import gspread
from handover_processing import operation_execute
from tkinter import messagebox
from oauth2client.service_account import ServiceAccountCredentials


def google_sheet_data(session):
    print('Fetching Google Sheet')

    # Setting up the API credentials
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]

    print("Defined scope")
    creds = ServiceAccountCredentials.from_json_keyfile_name("static/client_secret_969260230917-8lvqvqoq42uaobpsoqnfkdmc8sdrrbou.apps.googleusercontent.com.json", scope)
    try:
        client = gspread.authorize(creds)
    except Exception as e:
        messagebox.showerror('Connection Error', f"Failed to authorize or fetch data from Google: {e}")
        return

    # Open the desired spreadsheet (by its name in this case)
    try:
        sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1WPjBycSbDGcd1nTzQ6ZJ8xlTZU3z8uTDIThrCNqnrJU")
    except Exception as e:
        messagebox.showerror('Connection Error', f"Failed to fetch data from Google spreadsheet: {e}")
        return

    try:
        main_worksheet = sheet.worksheet("Form responses 1")
        key_account_worksheet = sheet.worksheet("Key Account Handover Requests")
        vendor_worksheet = sheet.worksheet("Vendors")
    except Exception as e:
        messagebox.showerror('Connection Error', f"Failed to fetch data from Google spreadsheet: {e}")
        return

    headers1 = main_worksheet.row_values(1)
    print("Headers for Form Responses:", headers1)
    dict_of_rows1 = main_worksheet.get_all_records()
    print('Form Responses tab records saved')

    headers2 = key_account_worksheet.row_values(1)
    print("Headers for Key Accounts:", headers2)

    # Replace empty headers with unique placeholders
    for i, header in enumerate(headers2):
        if header == '':
            headers2[i] = f"Placeholder_{i}"

    print("Headers after modification:", headers2)

    # Fetch all data without headers
    data = key_account_worksheet.get_all_values()[1:]  # Skip the header row
    # Combine headers with data rows
    dict_of_rows2 = [dict(zip(headers2, row)) for row in data]

    dict_of_vendors = vendor_worksheet.get_all_records()
    list_of_po_numbers = []

    headers1 = main_worksheet.row_values(1)
    headers2 = key_account_worksheet.row_values(1)

    try:
        po_number_column_number1 = headers1.index('PO number') + 1  # 1-based index
        po_number_column_number2 = headers2.index('PO number') + 1  # 1-based index
    except ValueError as e:
        messagebox.showerror('Column Error', f"There is no column \"PO number\" in the spreadsheet: {e}")
        return

    row_number = 1
    print('Fetching Form Responses tab')
    for row in dict_of_rows1:
        row_number += 1

        if row["Systems updated"] in ['FALSE', 'False', 0 , '0', '']:
            print(f"Row {row_number}: Form Responses tab - Systems Update is False...")
            continue

        if row["Handover Costs"] == '':
            print(f"Row {row_number}: Form Responses tab - Handover Costs is empty...")
            continue

        if row["PO number"] != '':
            print(f"Row {row_number}: Form Responses tab - already has the PO number...")
            continue

        sap_job_no_raw = row["SAP job no."]

        if sap_job_no_raw == '' or sap_job_no_raw is None:  # Check if it's empty or None and skip
            print(f'Row {row_number}: Form Responses tab - SAP Job No. is empty...')
            continue

        if isinstance(sap_job_no_raw, str) and any(char in sap_job_no_raw for char in ",;/"):
            print(f"f'Row {row_number}: Form Responses tab - multiple SAP job numbers detected for {sap_job_no_raw}...")
            continue

        print(f"Processing row {row_number} in Form Responses tab with SAP Job No. {row['SAP job no.']}...")

        if isinstance(sap_job_no_raw, int):  # Check if it's an integer
            sap_job_no = str(sap_job_no_raw)  # Convert to string
        else:
            sap_job_no = sap_job_no_raw.strip()  # Removing any leading/trailing whitespaces

        for vendor in dict_of_vendors:
            if vendor["ASC responsible"] == row["ASC responsible"]:
                asc_responsible = row["ASC responsible"]
                vendor_number = vendor["SAP vendor number"]
                break

        order_price = float(row["Handover Costs"])

        try:
            print(f"SAP job number: {sap_job_no}, ASC Responsible: {asc_responsible}, Vendor Number: {vendor_number}, Order Price: {order_price}")

            po_number = operation_execute(session, sap_job_no, vendor_number, order_price, asc_responsible)
            print(9)
            if po_number == "Blocked":
                pass
            else:
                main_worksheet.update_cell(row_number, po_number_column_number1, po_number)
                list_of_po_numbers.append(po_number)
        except Exception as e:
            print(f'Error: {e}')

    row_number = 1

    print('Fetching Key Account Handover Requests tab')
    for row in dict_of_rows2:
        row_number += 1

        if row["Systems updated"] in ['FALSE', 'False', 0 , '0', '']:
            print(f"Row {row_number}: Key Account tab - Systems Update is False...")
            continue

        if row["Handover Costs"] == '':
            print(f"Row {row_number}: Key Accounts tab - Handover Costs is empty...")
            continue

        if row["PO number"] != '':
            print(f"Row {row_number}: Key Accounts tab - already has the PO number...")
            continue

        sap_job_no_raw = row["SAP job no."]

        if sap_job_no_raw == '' or sap_job_no_raw is None:  # Check if it's empty or None and skip
            print(f'Row {row_number}: Key Accounts tab - SAP Job No. is empty...')
            continue

        if isinstance(sap_job_no_raw, str) and any(char in sap_job_no_raw for char in ",;/"):
            print(f"f'Row {row_number}: Key Accounts tab - multiple SAP job numbers detected for {sap_job_no_raw}...")
            continue

        print(f"Processing row {row_number} in Key Accounts tab with SAP Job No. {row['SAP job no.']}...")

        if isinstance(sap_job_no_raw, int):  # Check if it's an integer
            sap_job_no = str(sap_job_no_raw)  # Convert to string
        else:
            sap_job_no = sap_job_no_raw.strip()  # Removing any leading/trailing whitespaces

        for vendor in dict_of_vendors:
            if vendor["ASC responsible"] == row["ASC responsible"]:
                asc_responsible = row["ASC responsible"]
                vendor_number = vendor["SAP vendor number"]
                break

        order_price = float(row["Handover Costs"])

        try:
            print(f"SAP job number: {sap_job_no}, ASC Responsible: {asc_responsible}, Vendor Number: {vendor_number}, Order Price: {order_price}")

            po_number = operation_execute(session, sap_job_no, vendor_number, order_price, asc_responsible)
            if po_number == "Blocked":
                print("blocked")
            else:
                key_account_worksheet.update_cell(row_number, po_number_column_number2, po_number)
                list_of_po_numbers.append(po_number)
        except Exception as e:
            print(f'Error: {e}')

    if len(list_of_po_numbers):
        messagebox.showinfo(title="PO Numbers", message="PO numbers processed: " + ', '.join(map(str, list_of_po_numbers)))
    else:
        messagebox.showinfo(title="PO Numbers", message="No PO's to be processed, closing...")

#google_sheet_data('hi')