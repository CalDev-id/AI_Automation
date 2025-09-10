from googleapiclient.discovery import build
from google.oauth2 import service_account
import re


# --- AUTHENTICATION ---
def connect(doc_ref: str, json_path: str = "../projecttelkom-58f002bf8fa0.json", by: str = "id"):
    scope = [
        "https://www.googleapis.com/auth/documents",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = service_account.Credentials.from_service_account_file(json_path, scopes=scope)
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    if by == "id":
        return docs_service, doc_ref
    elif by == "url":
        match = re.search(r"/document/d/([a-zA-Z0-9-_]+)", doc_ref)
        if not match:
            raise ValueError("URL tidak valid untuk Google Docs")
        return docs_service, match.group(1)
    elif by == "name":
        results = drive_service.files().list(
            q=f"name='{doc_ref}' and mimeType='application/vnd.google-apps.document'",
            fields="files(id, name)"
        ).execute()
        files = results.get("files", [])
        if not files:
            raise ValueError(f"Tidak ada dokumen dengan nama '{doc_ref}'")
        return docs_service, files[0]["id"]
    else:
        raise ValueError("Param 'by' harus 'id', 'url', atau 'name'")


# --- READ ---
def read_doc(docs_service, doc_id):
    doc = docs_service.documents().get(documentId=doc_id).execute()
    content = doc.get("body").get("content")
    text = ""
    for element in content:
        if "paragraph" in element:
            for run in element["paragraph"]["elements"]:
                if "textRun" in run:
                    text += run["textRun"]["content"]
    print("#read_doc =>", text.strip())
    return text.strip()


# --- APPEND ---
def append_text(docs_service, doc_id, text):
    requests = [
        {
            "insertText": {
                "location": {"index": 1},
                "text": text + "\n"
            }
        }
    ]
    docs_service.documents().batchUpdate(
        documentId=doc_id, body={"requests": requests}
    ).execute()
    print("#append_text =>", text)


# --- UPDATE ---
def replace_text(docs_service, doc_id, old_text, new_text):
    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": old_text, "matchCase": True},
                "replaceText": new_text,
            }
        }
    ]
    docs_service.documents().batchUpdate(
        documentId=doc_id, body={"requests": requests}
    ).execute()
    print(f"#replace_text => '{old_text}' → '{new_text}'")


# --- CLEAR (hapus semua isi) ---
def clear_doc(docs_service, doc_id):
    doc = docs_service.documents().get(documentId=doc_id).execute()
    end_index = doc.get("body").get("content")[-1]["endIndex"]

    requests = [
        {
            "deleteContentRange": {
                "range": {"startIndex": 1, "endIndex": end_index - 1}
            }
        }
    ]
    docs_service.documents().batchUpdate(
        documentId=doc_id, body={"requests": requests}
    ).execute()
    print("#clear_doc => semua isi dihapus")


# --- CONTOH MAIN ---
if __name__ == "__main__":
    # Koneksi ke Docs
    # docs_service, doc_id = connect("Nama Dokumen", by="name")
    # docs_service, doc_id = connect("1xHvmZV-H8c-_6-p9dbsb_KadmDWUu4iooSkENgwuuXg", by="id")
    docs_service, doc_id = connect("https://docs.google.com/document/d/1135w_r0zCmDkSkabjip-WptNCdkicThfK5tG9TofBYU/edit?usp=sharing", by="url")

    # APPEND
    append_text(docs_service, doc_id, "Tambahan baru dari script Python!")

    # READ
    read_doc(docs_service, doc_id)

    # UPDATE
    replace_text(docs_service, doc_id, "Halo", "Hello")

    # CLEAR
    # clear_doc(docs_service, doc_id)
