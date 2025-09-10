import gspread
import pickle
import os

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import gspread
from gspread.exceptions import APIError


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_credentials(client_secret_path="../client_secret_425132504468-3v7vfinniqrlrevdb6sl3hivd8ulprd7.apps.googleusercontent.com.json", token_path="token.pkl"):
    creds = None
    if os.path.exists(token_path):
        with open(token_path, "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "wb") as token:
            pickle.dump(creds, token)

    return creds


def connect(sheet_ref: str, by: str = "name", client_secret_path="../client_secret_425132504468-3v7vfinniqrlrevdb6sl3hivd8ulprd7.apps.googleusercontent.com.json"):
    creds = get_credentials(client_secret_path)
    client = gspread.authorize(creds)

    if by == "name":
        return client.open(sheet_ref).sheet1
    elif by == "id":
        return client.open_by_key(sheet_ref).sheet1
    elif by == "url":
        return client.open_by_url(sheet_ref).sheet1
    else:
        raise ValueError("Param 'by' harus 'name', 'id', atau 'url'")


def create_spreadsheet(title: str, share_with: str = None, client_secret_path="../client_secret_425132504468-3v7vfinniqrlrevdb6sl3hivd8ulprd7.apps.googleusercontent.com.json"):
    creds = get_credentials(client_secret_path)
    client = gspread.authorize(creds)

    spreadsheet = client.create(title)
    print(f"#create spreadsheet => '{title}' dibuat dengan ID: {spreadsheet.id}")

    if share_with:
        spreadsheet.share(share_with, perm_type="user", role="writer")
        print(f"#share => dibagikan ke {share_with}")

    return spreadsheet


# --- CREATE / APPEND ---
def append_row(sheet, values: list):
    """Tambah row baru di akhir sheet."""
    sheet.append_row(values)
    print("#append_row =>", values)


# lalu di append_or_update pakai:
def append_or_update(sheet, key: str, values: list):
    """Cari berdasarkan key di kolom A. Jika ada update, kalau tidak append."""
    try:
        cell = sheet.find(key)
        sheet.update(range_name=f"A{cell.row}:Z{cell.row}", values=[values])
        print(f"#append_or_update_row => update {key} (row {cell.row})")
    except APIError:
        sheet.append_row(values)
        print(f"#append_or_update_row => append {key}")


def create_sheet(spreadsheet, title="Sheet Baru", rows=100, cols=20):
    """Buat sheet baru dalam spreadsheet."""
    # gunakan int untuk rows/cols
    new_sheet = spreadsheet.add_worksheet(title=title, rows=int(rows), cols=int(cols))
    print(f"#create => sheet {title} dibuat")
    return new_sheet


# --- READ ---
def get_row(sheet, row_num: int):
    row = sheet.row_values(row_num)
    print(f"#get row {row_num} =>", row)
    return row


def get_all(sheet):
    all_data = sheet.get_all_values()
    print("#get all rows =>", all_data)
    return all_data


def get_column(sheet, col_num: int):
    col = sheet.col_values(col_num)
    print(f"#get column {col_num} =>", col)
    return col


# --- UPDATE ---
def update_row(sheet, row_num: int, values: list):
    # gunakan named args utk menghindari peringatan deprecation
    sheet.update(range_name=f"A{row_num}:Z{row_num}", values=[values])
    print(f"#update row {row_num} =>", values)


def update_cell(sheet, cell: str, value):
    # update a single cell: gunakan update_acell agar semantik jelas
    sheet.update_acell(cell, value)
    print(f"#update cell {cell} => {value}")


# --- DELETE / CLEAR ---
def delete_row(sheet, row_num: int):
    sheet.delete_rows(row_num)
    print(f"#delete row {row_num}")


def delete_column(sheet, col_num: int):
    sheet.delete_columns(col_num)
    print(f"#delete column {col_num}")


def clear_sheet(sheet):
    sheet.clear()
    print("#clear => semua data dihapus")


def delete_sheet(spreadsheet, sheet_obj):
    spreadsheet.del_worksheet(sheet_obj)
    print(f"#delete sheet {sheet_obj.title}")


# ==============================
# DEMO
# ==============================
if __name__ == "__main__":
    # Buat spreadsheet baru (langsung muncul di Google Drive akun kamu)
    new_spreadsheet = create_spreadsheet("Spreadsheet Baru jirrrr", share_with="emailkamu@gmail.com")
    sheet = new_spreadsheet.sheet1

    # CREATE / APPEND
    append_row(sheet, ["Nama", "Alamat", "Email", "12345"])
    append_or_update(sheet, "Nama", ["Nama", "Alamat Baru", "Email Baru", "9999"])

    # READ
    get_row(sheet, 2)
    get_all(sheet)
    get_column(sheet, 1)

    # UPDATE
    update_row(sheet, 2, ["Update1", "Update2", "Update3", "Update4"])
    update_cell(sheet, "B2", "SatuSelUpdate")

    # DELETE
    delete_row(sheet, 2)
    delete_column(sheet, 3)
    # clear_sheet(sheet)  # hati-hati hapus semua isi
