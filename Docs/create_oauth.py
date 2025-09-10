import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# --- AUTHENTICATION ---
SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive"
]

def get_credentials():
    creds = None
    if os.path.exists("token_docs.pkl"):
        with open("token_docs.pkl", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # ganti path ini ke file OAuth Client ID JSON yang kamu download
            flow = InstalledAppFlow.from_client_secrets_file(
                "../client_secret_425132504468-3v7vfinniqrlrevdb6sl3hivd8ulprd7.apps.googleusercontent.com.json",
                SCOPES
            )
            creds = flow.run_local_server(port=0)

        # simpan supaya tidak login ulang
        with open("token_docs.pkl", "wb") as token:
            pickle.dump(creds, token)

    return creds


# --- INIT SERVICES ---
creds = get_credentials()
docs_service = build("docs", "v1", credentials=creds)
drive_service = build("drive", "v3", credentials=creds)


# --- CREATE ---
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
    print("#read =>", text.strip())
    return text.strip()


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
    docs_service.documents().batchUpdate(
        documentId=document_id, body={"requests": requests}
    ).execute()
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
    docs_service.documents().batchUpdate(
        documentId=document_id, body={"requests": requests}
    ).execute()
    print(f"#update => replaced '{old_text}' with '{new_text}'")


# --- DELETE DOCUMENT (via Drive API) ---
def delete_doc(document_id):
    drive_service.files().delete(fileId=document_id).execute()
    print("#delete => document deleted", document_id)


# --- MAIN DEMO ---
if __name__ == "__main__":
    # CREATE
    doc_id = create_doc("Test CRUD Docs OAuth", folder_id="1-jaWHdEjXhxYsgIgaCcPtgaBMEfd_MO-")

    # UPDATE (append)
    append_text(doc_id, "Halo, ini teks pertama.")
    append_text(doc_id, "Ini teks kedua.")

    # READ
    read_doc(doc_id)

    # UPDATE (replace)
    replace_text(doc_id, "Halo", "Hello")

    # DELETE
    # delete_doc(doc_id)   # hati-hati, hapus permanen
