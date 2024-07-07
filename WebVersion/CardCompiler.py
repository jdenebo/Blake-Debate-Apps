import os
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from docx import Document
from docxcompose.composer import Composer
from docxcompose.composer import Composer
from docx import Document as Document_compose

# Set up Google Drive API credentials
SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'creds.json'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

service = build('drive', 'v3', credentials=credentials)

# Specify the folder ID
folder_id = ''

# List all .docx files in the folder
results = service.files().list(
    q=f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'",
    spaces='drive',
    fields='files(id, name)').execute()
items = results.get('files', [])

# Function to download a file
def download_file(file_id, file_name):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(f"Download {file_name}: {int(status.progress() * 100)}%.")
    file_bytes = fh.getvalue()
    return file_bytes

def move_file(file_id):
    # Retrieve the existing parents to remove
    file = service.files().get(fileId=file_id,
                               fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    # Move the file to the new folder
    file = service.files().update(fileId=file_id,
                                addParents="1M2kSWXSGXF2Sw_9G5kAMw6z0Yh9PmYW9",
                                removeParents=previous_parents,
                                fields='id, parents').execute()

def move_pdf(file_id, new_folder_id):
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    file = service.files().update(fileId=file_id,
                                  addParents=new_folder_id,
                                  removeParents=previous_parents,
                                  fields='id, parents').execute()

def list_files_in_folder(folder_id):
    query = f"'{folder_id}' in parents and mimeType='application/pdf'"
    results = service.files().list(q=query, pageSize=100, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])
    return items

def move_pdf_files_to_subfolder():
    pdf_files = list_files_in_folder("1LrXZWRoh7NGsK8sRqcmBTJoGTd3VuLWW")
    subfolder_id = "1im5TtBxricq31TNtCCxRhu73aZqDgLuh"
    
    for file in pdf_files:
        move_pdf(file['id'], subfolder_id)
        print(f"Moved file: {file['name']} to  PDF")


# Download all .docx files
docx_files = []
with open ("card-comp.docx", "a") as f:
    with open ("new-card.docx", "wb") as f2:
        for item in items:
            file_id = item['id']
            file_name = item['name']
            bts = download_file(file_id, file_name)
            docx_files.append(file_name)
            move_file(file_id)
            f2.write(bts)
            composer = Composer(Document_compose(os.path.abspath("card-comp.docx")))
            composer.append(Document_compose("new-card.docx"))
            composer.save("card-comp.docx")

move_pdf_files_to_subfolder()
