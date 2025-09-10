from googleapiclient.discovery import build
from google.oauth2 import service_account

# --- AUTHENTICATION ---
SCOPES = ["https://www.googleapis.com/auth/documents", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = "../projecttelkom-58f002bf8fa0.json"

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

docs_service = build("docs", "v1", credentials=creds)
drive_service = build("drive", "v3", credentials=creds)


# --- CREATE ---
# def create_doc(title="Dokumen Baru"):
#     body = {"title": title}
#     doc = docs_service.documents().create(body=body).execute()
#     print("#create =>", f"Created document with ID: {doc['documentId']}")
#     return doc["documentId"]
def create_doc(title="Dokumen Baru", folder_id=None):
    file_metadata = {
        "name": title,
        "mimeType": "application/vnd.google-apps.document"
    }

    if folder_id:
        file_metadata["parents"] = [folder_id]

    file = drive_service.files().create(body=file_metadata).execute()
    print("#create => Created document with ID:", file["id"])
    return file["id"]



# --- READ ---
def read_doc(document_id):
    doc = docs_service.documents().get(documentId=document_id).execute()
    content = doc.get("body").get("content")
    text = ""
    for element in content:
        if "paragraph" in element:
            for run in element["paragraph"]["elements"]:
                if "textRun" in run:
                    text += run["textRun"]["content"]
    print("#read =>", text)
    return text


# --- UPDATE (append text) ---
def append_text(document_id, text):
    requests = [
        {
            "insertText": {
                "location": {"index": 1},
                "text": text + "\n"
            }
        }
    ]
    docs_service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()
    print("#update => appended text:", text)


# --- UPDATE (replace text) ---
def replace_text(document_id, old_text, new_text):
    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": old_text, "matchCase": True},
                "replaceText": new_text,
            }
        }
    ]
    docs_service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()
    print(f"#update => replaced '{old_text}' with '{new_text}'")


# --- DELETE DOCUMENT (via Drive API) ---
def delete_doc(document_id):
    drive_service.files().delete(fileId=document_id).execute()
    print("#delete => document deleted", document_id)


# --- MAIN DEMO ---
if __name__ == "__main__":
    # CREATE
    doc_id = create_doc("Test CRUD Docs", folder_id="1-jaWHdEjXhxYsgIgaCcPtgaBMEfd_MO-")

    # # UPDATE (append)
    # append_text(doc_id, "Halo, ini teks pertama.")
    # append_text(doc_id, "Ini teks kedua.")

    # # READ
    # read_doc(doc_id)

    # # UPDATE (replace)
    # replace_text(doc_id, "Halo", "Hello")

    # DELETE
    # delete_doc(doc_id)   # hati-hati, hapus permanen
