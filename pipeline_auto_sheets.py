import os
import json
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google import genai
from google.genai import types

# 1. API Authorization
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

service_account_info = json.loads(os.environ["GDRIVE_SERVICE_ACCOUNT_JSON"])
creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)

drive_service = build("drive", "v3", credentials=creds)
gc = gspread.authorize(creds)
ai_client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

INVOICES_FOLDER_ID = os.environ["INVOICES_FOLDER_ID"]
SHEET_TITLE = "AP_CFO_Executive_Intelligence"

def get_or_create_spreadsheet():
    """Finds existing spreadsheet inside the folder or creates a new configured one."""
    query = f"'{INVOICES_FOLDER_ID}' in parents and name='{SHEET_TITLE}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name, webViewLink)").execute()
    files = results.get("files", [])

    if files:
        print(f"Connecting to existing spreadsheet: {files[0]['webViewLink']}")
        return gc.open_by_key(files[0]["id"]), files[0]["webViewLink"]

    print("Spreadsheet not found. Creating a new one with automated CFO dashboard...")
    file_metadata = {
        "name": SHEET_TITLE,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [INVOICES_FOLDER_ID]
    }
    new_file = drive_service.files().create(body=file_metadata, fields="id, webViewLink").execute()
    sh = gc.open_by_key(new_file["id"])

    # 1. Setup Invoice_Headers sheet
    default_sheet = sh.sheet1
    default_sheet.update_title("Invoice_Headers")
    headers = [
        "Invoice ID", "Date Processed", "Vendor Name", "Invoice Number",
        "Invoice Date", "Due Date", "Currency", "Subtotal", "Tax Amount",
        "Grand Total", "Status", "Drive File Link"
    ]
    default_sheet.append_row(headers, value_input_option="USER_ENTERED")

    # 2. Setup Line_Items sheet
    items_sheet = sh.add_worksheet(title="Line_Items", rows=1000, cols=6)
    item_headers = ["Invoice Number", "Vendor Name", "Description", "Quantity", "Unit Price", "Total Price"]
    items_sheet.append_row(item_headers, value_input_option="USER_ENTERED")

    # 3. Setup CFO_Dashboard sheet with KPI formulas
    dash_sheet = sh.add_worksheet(title="CFO_Dashboard", rows=100, cols=12)
    dash_setup = [
        ["CFO AP EXECUTIVE INTELLIGENCE DASHBOARD", ""],
        ["", ""],
        ["KPI Metric", "Value"],
        ["Total Invoices Processed", "=COUNTA(Invoice_Headers!A2:A)"],
        ["Total Spend (USD)", '=SUMIF(Invoice_Headers!G2:G, "USD", Invoice_Headers!J2:J)'],
        ["Total Spend (BDT)", '=SUMIF(Invoice_Headers!G2:G, "BDT", Invoice_Headers!J2:J)'],
        ["Total Spend (EUR)", '=SUMIF(Invoice_Headers!G2:G, "EUR", Invoice_Headers!J2:J)'],
        ["Total Spend (GBP)", '=SUMIF(Invoice_Headers!G2:G, "GBP", Invoice_Headers!J2:J)'],
        ["", ""],
        ["TOP 5 VENDORS BY TOTAL SPEND", ""],
        ['=QUERY(Invoice_Headers!C2:J, "SELECT C, SUM(J) WHERE C IS NOT NULL GROUP BY C ORDER BY SUM(J) DESC LIMIT 5 LABEL C \'Vendor\', SUM(J) \'Total Spend\'")', ""]
    ]
    dash_sheet.update("A1:B11", dash_setup, value_input_option="USER_ENTERED")

    print(f"Spreadsheet initialized successfully: {new_file['webViewLink']}")
    return sh, new_file["webViewLink"]

def main():
    sh, sheet_url = get_or_create_spreadsheet()
    header_sheet = sh.worksheet("Invoice_Headers")
    items_sheet = sh.worksheet("Line_Items")

    # Query all PDF invoices in the folder
    q = f"'{INVOICES_FOLDER_ID}' in parents and mimeType='application/pdf' and trashed=false"
    results = drive_service.files().list(q=q, fields="files(id, name, webViewLink, createdTime)", pageSize=100).execute()
    files = results.get("files", [])

    if not files:
        print("No invoice PDFs found in the folder.")
        return

    # Check already processed links (Column 12)
    col_links = header_sheet.col_values(12)
    processed_links = set(col_links[1:]) if len(col_links) > 1 else set()

    unprocessed_files = [f for f in files if f["webViewLink"] not in processed_links]

    if not unprocessed_files:
        print("All invoice files are already extracted and synced.")
        return

    print(f"Processing {len(unprocessed_files)} new invoices...")

    prompt = """
    Extract all details from this invoice accurately into this exact JSON structure:
    {
      "vendor_name": "string",
      "invoice_number": "string",
      "invoice_date": "YYYY-MM-DD",
      "due_date": "YYYY-MM-DD",
      "currency": "USD/BDT/EUR/GBP/CAD/AUD/SGD",
      "subtotal": 0.0,
      "tax_amount": 0.0,
      "grand_total": 0.0,
      "line_items": [
        {
          "description": "string",
          "quantity": 1,
          "unit_price": 0.0,
          "total_price": 0.0
        }
      ]
    }
    Return ONLY valid raw JSON without markdown or backticks.
    """

    for file in unprocessed_files:
        file_id = file["id"]
        file_name = file["name"]
        file_link = file["webViewLink"]
        print(f"Analyzing {file_name}...")

        try:
            pdf_bytes = drive_service.files().get_media(fileId=file_id).execute()

            response = ai_client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[
                    types.Part.from_bytes(data=pdf_bytes, mime_type='application/pdf'),
                    prompt
                ],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )

            data = json.loads(response.text.strip())
            inv_num = data.get("invoice_number") or file_name
            vendor = data.get("vendor_name") or "Unknown"

            # 1. Log Invoice Header Row
            header_sheet.append_row([
                inv_num,
                file.get("createdTime", "")[:10],
                vendor,
                inv_num,
                data.get("invoice_date", ""),
                data.get("due_date", ""),
                data.get("currency", "USD"),
                float(data.get("subtotal") or 0.0),
                float(data.get("tax_amount") or 0.0),
                float(data.get("grand_total") or 0.0),
                "Verified",
                file_link
            ], value_input_option="USER_ENTERED")

            # 2. Log Line Items
            line_rows = []
            for item in data.get("line_items", []):
                line_rows.append([
                    inv_num,
                    vendor,
                    item.get("description", ""),
                    float(item.get("quantity") or 1),
                    float(item.get("unit_price") or 0.0),
                    float(item.get("total_price") or 0.0)
                ])

            if line_rows:
                items_sheet.append_rows(line_rows, value_input_option="USER_ENTERED")

            print(f"Logged: {inv_num} | {vendor}")

        except Exception as e:
            print(f"Failed to process {file_name}: {str(e)}")

    # Write output link for GitHub Actions / Make.com Webhook
    with open("sheet_info.json", "w") as f:
        json.dump({"sheet_url": sheet_url, "processed_count": len(unprocessed_files)}, f)

    print(f"\nPipeline Complete. Live Dashboard: {sheet_url}")

if __name__ == "__main__":
    main()
