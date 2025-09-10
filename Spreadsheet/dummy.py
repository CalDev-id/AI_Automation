import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name("../projecttelkom-58f002bf8fa0.json", scope)

client = gspread.authorize(credentials)

sheet = client.open("Template Input").sheet1

first_row = sheet.row_values(2)

print(first_row)
