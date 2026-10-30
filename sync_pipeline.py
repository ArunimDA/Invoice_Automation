import os
import io
import json
from datetime import datetime
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import gspread
from google import genai

# Secrets loaded from GitHub Action environment variables
GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]
FOLDER_ID = os.environ["INVOICES_FOLDER_ID"]
SA_JSON_STRING = os.environ["GDRIVE_SERVICE_ACCOUNT_JSON"]
MAKE_WEBHOOK_URL = os.environ.get("MAKE_CFO_WEBHOOK_URL", "")

SPREADSHEET_NAME = "AP_CFO_Executive_Intelligence"
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

def get_clients():
    creds_dict = json.loads(SA_JSON_STRING)
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    drive_service = build("drive", "v3", credentials=creds)
    gc = gspread.authorize(creds)
    ai_client = genai.Client(api_key=GEMINI_API_KEY)
    return drive_service, gc, ai_client

def get_or_create_sheet(drive_service, gc):
    query = (
        f"name = '{SPREADSHEET_NAME}' and '{FOLDER_ID}' in parents "
        f"and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    )
    res = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = res.get("files", [])

    if files:
        return gc.open_by_key(files[0]["id"])

    metadata = {
        "name": SPREADSHEET_NAME,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [FOLDER_ID]
    }
    file = drive_service.files().create(body=metadata, fields="id").execute()
    sh = gc.open_by_key(file["id"])

    # 1. Invoice_Headers
    ws_headers = sh.sheet1
    ws_headers.update_title("Invoice_Headers")
    ws_headers.append_row([
        "Invoice ID", "Date Processed", "Vendor Name", "Invoice Number",
        "Invoice Date", "Due Date", "Currency", "Subtotal", "Tax Amount",
        "Grand Total", "Status", "Drive File Link", "File ID"
    ])

    # 2. Line_Items
    ws_items = sh.add_worksheet(title="Line_Items", rows=1000, cols=10)
    ws_items.append_row([
        "Invoice Number", "Vendor Name", "Description", "Quantity", "Unit Price", "Total Price", "File ID"
    ])

    # 3. CFO Dashboard Formulas
    ws_dash = sh.add_worksheet(title="CFO_Dashboard", rows=50, cols=10)
    ws_dash.update(
        range_name="A1:B11",
        values=[
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
        ],
        value_input_option="USER_ENTERED"
    )
    return sh

def extract_pdf_data(client, pdf_bytes):
    prompt = """Extract invoice details into strict valid JSON format:
{
  "vendor_name": "string",
  "invoice_number": "string",
  "invoice_date": "YYYY-MM-DD",
  "due_date": "YYYY-MM-DD",
  "currency": "USD",
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
Return ONLY pure JSON. No markdown ticks."""

    response = client.models.generate_content(
        model="gemini-1.5-flash",
        contents=[
            genai.types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            prompt
        ],
        config=genai.types.GenerateContentConfig(response_mime_type="application/json")
    )
    text = response.text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)

def main():
    drive_service, gc, ai_client = get_clients()
    sh = get_or_create_sheet(drive_service, gc)
    ws_headers = sh.worksheet("Invoice_Headers")
    ws_items = sh.worksheet("Line_Items")

    # Fetch live files in Google Drive folder
    query = f"'{FOLDER_ID}' in parents and mimeType = 'application/pdf' and trashed = false"
    res = drive_service.files().list(q=query, fields="files(id, name, webViewLink)").execute()
    live_files = {f["id"]: f for f in res.get("files", [])}

    # Fetch current state of Google Sheet
    header_rows = ws_headers.get_all_values()
    sheet_file_ids = {}
    if len(header_rows) > 1:
        for idx, row in enumerate(header_rows[1:], start=2):
            if len(row) >= 13 and row[12]:
                sheet_file_ids[row[12]] = idx

    changes_detected = False

    # 1. Handle Deletions (Files removed from Drive but still present in Sheet)
    deleted_ids = [fid for fid in sheet_file_ids.keys() if fid not in live_files]
    if deleted_ids:
        print(f"Detected {len(deleted_ids)} deleted files. Purging from ledger...")
        # Clean line items
        all_item_rows = ws_items.get_all_values()
        if len(all_item_rows) > 1:
            filtered_items = [all_item_rows[0]] + [r for r in all_item_rows[1:] if len(r) >= 7 and r[6] not in deleted_ids]
            ws_items.clear()
            ws_items.update("A1", filtered_items)

        # Clean headers
        filtered_headers = [header_rows[0]] + [r for r in header_rows[1:] if len(r) >= 13 and r[12] not in deleted_ids]
        ws_headers.clear()
        ws_headers.update("A1", filtered_headers)
        changes_detected = True

    # 2. Handle Additions (New files in Drive not yet processed)
    new_files = [f for fid, f in live_files.items() if fid not in sheet_file_ids]
    print(f"Detected {len(new_files)} new files to ingest.")

    for f in new_files:
        print(f"Processing: {f['name']}...")
        req = drive_service.files().get_media(fileId=f["id"])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        try:
            data = extract_pdf_data(ai_client, fh.getvalue())
            inv_num = data.get("invoice_number") or f["name"]
            vendor = data.get("vendor_name") or "Unknown"

            ws_headers.append_row([
                inv_num,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                vendor,
                inv_num,
                data.get("invoice_date", ""),
                data.get("due_date", ""),
                data.get("currency", "USD"),
                data.get("subtotal", 0),
                data.get("tax_amount", 0),
                data.get("grand_total", 0),
                "Verified",
                f.get("webViewLink", ""),
                f["id"]
            ])

            line_items = data.get("line_items", [])
            if line_items:
                items_batch = [
                    [inv_num, vendor, item.get("description", ""), item.get("quantity", 1), item.get("unit_price", 0), item.get("total_price", 0), f["id"]]
                    for item in line_items
                ]
                ws_items.append_rows(items_batch)

            changes_detected = True
        except Exception as err:
            print(f"Error processing {f['name']}: {err}")

    # 3. Notify Make.com to email CFO if changes were processed
    if changes_detected and MAKE_WEBHOOK_URL:
        dash_vals = sh.worksheet("CFO_Dashboard").get_all_values()
        total_invoices = dash_vals[3][1] if len(dash_vals) > 3 else "N/A"
        usd_spend = dash_vals[4][1] if len(dash_vals) > 4 else "0"

        payload = {
            "status": "success",
            "spreadsheet_url": sh.url,
            "total_invoices": total_invoices,
            "total_usd_spend": usd_spend,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        requests.post(MAKE_WEBHOOK_URL, json=payload)
        print("CFO notification webhook triggered successfully.")

if __name__ == "__main__":
    main()
