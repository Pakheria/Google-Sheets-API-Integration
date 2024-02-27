import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *
import warnings

warnings.filterwarnings("ignore", message=".*worksheet.update.*")


scopes = [ "https://www.googleapis.com/auth/spreadsheets" ]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

sheet_id = "1K1a-xefRA1yaaMAuHWMz7-LupYZWQ408EEvUeK0oetg"
workbook = client.open_by_key(sheet_id)

values = [
    ["Name", "Price", "Quantity"],
    ["Basketball", 29.99, 1],
    ["Jeans", 39.99, 4],
    ["Soap", 7.99, 3],
]


# add new worksheet if not avaialble
sheets = map(lambda x: x.title, workbook.worksheets())
new_worksheet_name = "Values"

if new_worksheet_name in sheets:
    sheet = workbook.worksheet(new_worksheet_name)
else:
    sheet = workbook.add_worksheet(new_worksheet_name, rows=10, cols=10)
    
sheet.clear()

sheet.update(values, f"A1:C{len(values)}")
sheet.update_cell(len(values) + 1, 2, "=SUM(B2:B4)")
sheet.update_cell(len(values) + 1, 3, "=SUM(C2:C4)")

# sheet.cell["A1:C1"].font.bold = True
sheet.format("A1:C1", {"textFormat" : {"bold" : True}})