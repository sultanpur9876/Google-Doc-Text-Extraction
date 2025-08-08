import os
import io
import subprocess
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# If modifying these scopes, delete the token.json
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds

def download_doc_as_docx(doc_id, output_path='document.docx'):
    creds = authenticate()
    service = build('drive', 'v3', credentials=creds)

    request = service.files().export_media(fileId=doc_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    fh = io.FileIO(output_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(f"Download progress: {int(status.progress() * 100)}%")

    print(f"Downloaded .docx to {output_path}")
    return output_path

def convert_docx_to_html(docx_path, html_output='output.html', css_path=None):
    media_dir = "media"
    os.makedirs(media_dir, exist_ok=True)

    cmd = [
        "pandoc", docx_path,
        "-f", "docx+table",
        "-t", "html",
        "-s", "-o", html_output,
        f"--extract-media={media_dir}"
    ]

    if css_path:
        cmd.append(f"--css={css_path}")

    subprocess.run(cmd, check=True)
    print(f"Converted HTML saved to {html_output}")
    return html_output

if __name__ == "__main__":
    # Replace this with your actual Google Doc ID
    GOOGLE_DOC_ID = "157u5ksxCDSdBsx9EC0dDCd3d0jAxOA5wJQc8QP_sFlI"

    docx_file = download_doc_as_docx(GOOGLE_DOC_ID)
    convert_docx_to_html(docx_file, css_path="style.css")  # Optional CSS
