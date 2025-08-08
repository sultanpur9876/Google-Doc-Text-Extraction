# from googleapiclient.discovery import build
# from google.oauth2.credentials import Credentials
# import io
# from googleapiclient.http import MediaIoBaseDownload

# creds = Credentials.from_authorized_user_file("token.json")
# drive_service = build("drive", "v3", credentials=creds)
 
# file_id = "157u5ksxCDSdBsx9EC0dDCd3d0jAxOA5wJQc8QP_sFlI"  # Replace with your doc ID
# request = drive_service.files().export_media(fileId=file_id, mimeType="text/html")

# fh = io.FileIO("output_drive.html", "wb")
# downloader = MediaIoBaseDownload(fh, request)
# done = False
# while not done:
#     status, done = downloader.next_chunk()
#     print(f"Download {int(status.progress() * 100)}%.")

# from bs4 import BeautifulSoup

# with open("output_drive.html", "r", encoding="utf-8") as f:
#     soup = BeautifulSoup(f, "html.parser")

# # Optional: remove unnecessary divs, styles, etc.
# for tag in soup.find_all(True):
#     tag.attrs = {}  # remove all attributes like style, class, etc.

# clean_html = soup.prettify()

# with open("cleaned_output.html", "w", encoding="utf-8") as f:
#     f.write(clean_html)

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.credentials import Credentials
import io

# Load credentials
creds = Credentials.from_authorized_user_file("token.json")

# Build the Drive API client
drive_service = build("drive", "v3", credentials=creds)

# Your Google Doc ID
file_id = "157u5ksxCDSdBsx9EC0dDCd3d0jAxOA5wJQc8QP_sFlI"


request = drive_service.files().export_media(
    fileId=file_id,
    mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

with open("doc_output.docx", "wb") as f:
    downloader = MediaIoBaseDownload(f, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()


import mammoth

# Convert .docx to clean HTML
with open("doc_output.docx", "rb") as docx_file:
    result = mammoth.convert_to_html(docx_file)
    html = result.value  # This is the clean HTML

# Add basic HTML structure if needed
full_html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Converted Google Doc</title>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
        h1, h2, h3 {{ color: #333; }}
        p {{ margin-bottom: 1em; }}
    </style>
</head>
<body>
{html}
</body>
</html>
"""

# Write to a .html file
with open("converted_doc.html", "w", encoding="utf-8") as file:
    file.write(full_html)

print("✅ HTML saved as 'converted_doc.html'")


# # Export the document as HTML
# request = drive_service.files().export_media(fileId=file_id, mimeType="text/html")

# # Option 1: Save as file (readable in browser)
# with open("document.html", "wb") as f:
#     downloader = MediaIoBaseDownload(f, request)
#     done = False
#     while not done:
#         status, done = downloader.next_chunk()
#     print("✅ HTML saved as document.html")

# # Option 2: Keep in memory as a string (for web apps / processing)
# fh = io.BytesIO()
# request = drive_service.files().export_media(fileId=file_id, mimeType="text/html")
# downloader = MediaIoBaseDownload(fh, request)
# done = False
# while not done:
#     status, done = downloader.next_chunk()

# # Decode to string
# html_string = fh.getvalue().decode("utf-8")
# print("✅ HTML string ready to use")
# # You can now return this from a web endpoint, render in a Streamlit app, etc.
