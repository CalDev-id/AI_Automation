import gspread
from oauth2client.service_account import ServiceAccountCredentials


# --- AUTHENTICATION ---
def connect(sheet_ref: str, json_path: str = "../projecttelkom-58f002bf8fa0.json", by: str = "name"):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
    client = gspread.authorize(credentials)

    if by == "name":
        return client.open(sheet_ref).sheet1
    elif by == "id":
        return client.open_by_key(sheet_ref).sheet1
    elif by == "url":
        return client.open_by_url(sheet_ref).sheet1
    else:
        raise ValueError("Param 'by' harus 'name', 'id', atau 'url'")



# --- CREATE / APPEND ---
def append_row(sheet, values: list):
    """Tambah row baru di akhir sheet."""
    sheet.append_row(values)
    print("#append_row =>", values)


def append_or_update(sheet, key: str, values: list):
    """Cari berdasarkan key di kolom A. Jika ada update, kalau tidak append."""
    cell = sheet.find(key)
    if cell:
        sheet.update(f"A{cell.row}:Z{cell.row}", [values])
        print(f"#append_or_update_row => update {key}")
    else:
        sheet.append_row(values)
        print(f"#append_or_update_row => append {key}")


def create_sheet(spreadsheet, title="Sheet Baru", rows=100, cols=20):
    """Buat sheet baru dalam spreadsheet."""
    new_sheet = spreadsheet.add_worksheet(title=title, rows=str(rows), cols=str(cols))
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
    sheet.update(f"A{row_num}:Z{row_num}", [values])
    print(f"#update row {row_num} =>", values)


def update_cell(sheet, cell: str, value):
    sheet.update(range_name=cell, values=[[value]])
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


# --- CONTOH MAIN ---
if __name__ == "__main__":
    # Koneksi ke sheet
    # pakai nama
    # sheet = connect("Template Input", by="name")

    # # pakai id
    # sheet = connect("1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890", by="id")

    # # pakai url
    sheet = connect("https://docs.google.com/spreadsheets/d/1R_NkbuYOQiAsNJu0hwx-MeS6dXJefQ7V-q6tBZVUFZo/edit?usp=sharing", by="url")

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

    # DELETE / CLEAR
    delete_row(sheet, 2)
    delete_column(sheet, 3)
    # clear_sheet(sheet)  # hati-hati, hapus semua isi
