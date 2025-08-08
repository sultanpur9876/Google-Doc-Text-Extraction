from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import os.path

# SCOPES = ["https://www.googleapis.com/auth/documents.readonly"]
# DOCUMENT_ID = "157u5ksxCDSdBsx9EC0dDCd3d0jAxOA5wJQc8QP_sFlI"
# DOCUMENT_ID = "1J66UPR8WpjzTwiTNIU1CXUUzIDELeRs6p9hwF75FBXk"  #Oranage Book
DOCUMENT_ID = "19u6bLSHgCx4dKBMkIBXt2-lIAEUrHUq-IZPQxDDOytY"   #Red Book
SCOPES = [
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

def extract_text_from_elements(elements):
    text_html = ""
    for el in elements:
        # print(el)
        if "textRun" in el:

            run = el["textRun"]
            content = run.get("content", "").strip()
            # print(content)
            if not content:
                continue

            style = run.get("textStyle", {})
            link = style.get("link", {}).get("url")
            fg_color = style.get("foregroundColor", {}).get("color", {}).get("rgbColor")
            # bg_color = style.get("backgroundColor", {}).get("color", {}).get("rgbColor")
            font_size = style.get("fontSize", {}).get("magnitude")
            font_unit = style.get("fontSize", {}).get("unit", "PT")
            font_family = style.get("weightedFontFamily", {}).get("fontFamily")
            font_weight = style.get("weightedFontFamily", {}).get("weight", 400)

            is_bold = style.get("bold", False)
            is_italic = style.get("italic", False)
            is_underline = style.get("underline", False)

            html_style = []
            if fg_color:
                r = int(fg_color.get("red", 0) * 255)
                g = int(fg_color.get("green", 0) * 255)
                b = int(fg_color.get("blue", 0) * 255)
                html_style.append(f"color: rgb({r},{g},{b});")

            # if bg_color:
            #     r = int(bg_color.get("red", 0) * 255)
            #     g = int(bg_color.get("green", 0) * 255)
            #     b = int(bg_color.get("blue", 0) * 255)
            #     html_style.append(f"background-color: rgb({r},{g},{b});")
            
            if font_size:
                html_style.append(f"font-size: {font_size}{font_unit.lower()};")

            # Add font family
            if font_family:
                html_style.append(f"font-family: '{font_family}', sans-serif;")
            
            if is_bold:
                html_style.append("font-weight: bold;")

            elif font_weight:
                html_style.append(f"font-weight: {font_weight};")

            if is_italic:
                html_style.append("font-style: italic;")

            if is_underline:
                html_style.append("text-decoration: underline;")

            html = content
            if html_style:
                html = f"<span style=\"{' '.join(html_style)}\">{html}</span>"
            if link:
                html = f"<a href=\"{link}\">{html}</a>"

            text_html += html

    # return text_html.strip(), content
    if text_html:
        combined_text = ''.join([el.get("textRun", {}).get("content", "") for el in elements if "textRun" in el])
        return text_html.strip(), combined_text.strip()
    else:
        return "", ""



def process_table(table):
    html = "<table border='1' style='border-collapse: collapse;'>\n"
    for row in table.get("tableRows", []):
        html += "  <tr>\n"
        for cell in row.get("tableCells", []):
            cell_text_html = ""
            for content in cell.get("content", []):
                if "paragraph" in content:
                    para = content["paragraph"]
                    elements = para.get("elements", [])
                    bullet = "bullet" in para
                    text_html, _ = extract_text_from_elements(elements)

                    # Get text alignment
                    para_style = para.get("paragraphStyle", {})
                    alignment = para_style.get("alignment", "START")
                    align_map = {
                        "START": "left",
                        "CENTER": "center",
                        "END": "right",
                        "JUSTIFIED": "justify"
                    }
                    css_align = align_map.get(alignment, "left")

                    if bullet:
                        cell_text_html += f"<ul><li>{text_html}</li></ul>"
                    else:
                        cell_text_html += f"<p style='text-align: {css_align};'>{text_html}</p>"

            # Extract background color
            cell_style = cell.get("tableCellStyle", {})
            # background = cell_style.get("backgroundColor", {}).get("color", {}).get("rgbColor", {})
            # bg_color = ""
            # if background:
            #     r = int(background.get("red", 1.0) * 255)
            #     g = int(background.get("green", 1.0) * 255)
            #     b = int(background.get("blue", 1.0) * 255)
            #     bg_color = f"background-color: rgb({r}, {g}, {b});"

            # Width and height (in points, convert to px)
            width_pt = cell_style.get("width", {}).get("magnitude")
            height_pt = cell_style.get("height", {}).get("magnitude")
            width_px = f"{width_pt * (96/72)}px;" if width_pt else ""
            height_px = f"{height_pt * (96/72)}px;" if height_pt else ""

            # Combine all styles
            # style = f" style='{bg_color}{width_px}{height_px}'".strip() if (bg_color or width_px or height_px) else ""

            # html += f" <td {style}>{cell_text_html} </td>\n"
            html += f" <td>{cell_text_html} </td>\n"

        html += "  </tr>\n"
    html += "</table>\n"
    return html


def process_image(element, inline_objects):
    image_id = element.get("inlineObjectElement", {}).get("inlineObjectId")
    if image_id and image_id in inline_objects:
        embedded_object = inline_objects[image_id]["inlineObjectProperties"]["embeddedObject"]
        source = embedded_object.get("imageProperties", {}).get("contentUri")
        title = embedded_object.get("title", "Image")
        if source:
            return f"<img src='{source}' alt='{title}' style='max-width:100%;'/>\n"
    return ""


def extract_content_as_html(document, target_heading):
    content_html = ""
    capture = False
    base_level = None
    inline_objects = document.get("inlineObjects", {})
    bullet_open = False
    # print("document = ",document)
    # output_file = "extracted_content.txt"
    # with open(output_file, "w", encoding="utf-8") as f:
    #     f.write(str(document))

    # print(f"Content has been written to {output_file}")
    for element in document.get("body").get("content", []):

        if "paragraph" in element:
            para = element["paragraph"]
            style_type = para.get("paragraphStyle", {}).get("namedStyleType", "")
            elements = para.get("elements", [])
            bullet = "bullet" in para

            text_html, text = extract_text_from_elements(elements)

            if style_type.startswith("HEADING_") and text_html:
                current_level = int(style_type.replace("HEADING_", ""))
                print("Text = ",text, "Length = ", len(text))
                if text.strip() == target_heading.strip():
                    print("✅ Match found for heading:", repr(text.strip()))

                    capture = True
                    base_level = current_level
                    continue
                if capture and current_level <= base_level:
                    if bullet_open:
                        content_html += "</ul>\n"
                        bullet_open = False
                    break
                if capture:
                    if bullet_open:
                        content_html += "</ul>\n"
                        bullet_open = False
                    content_html += f"<h{current_level}>{text_html}</h{current_level}>\n"
                continue

            if capture:
                if bullet:
                    if not bullet_open:
                        content_html += "<ul>\n"
                        bullet_open = True
                    content_html += f"<li>{text_html}</li>\n"
                else:
                    if bullet_open:
                        content_html += "</ul>\n"
                        bullet_open = False
                    content_html += f"<p>{text_html}</p>\n"

        elif "table" in element and capture:
            if bullet_open:
                content_html += "</ul>\n"
                bullet_open = False
            content_html += process_table(element["table"])

        elif "inlineObjectElement" in element and capture:
            if bullet_open:
                content_html += "</ul>\n"
                bullet_open = False
            content_html += process_image(element, inline_objects)

    if bullet_open:
        content_html += "</ul>\n"
    print("content_html = ", content_html)
    return content_html.strip()


def get_headings(document):
    headings = []
    for element in document.get("body").get("content", []):
        if "paragraph" in element:
            para = element["paragraph"]
            if "paragraphStyle" in para and "namedStyleType" in para["paragraphStyle"]:
                style = para["paragraphStyle"]["namedStyleType"]
                if style.startswith("HEADING_"):
                    text = ""
                    for el in para.get("elements", []):
                        text += el.get("textRun", {}).get("content", "")
                    headings.append((style, text.strip()))
    return headings


def main():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("docs", "v1", credentials=creds)
        document = service.documents().get(documentId=DOCUMENT_ID).execute()
        print(f"The title of the document is: {document.get('title')}")

        target = "Employee Salary"  # Change to the section heading you want
        html_output = extract_content_as_html(document, target)
        doc_style = document.get("documentStyle", {})
        page_width_pt = doc_style.get("pageSize", {}).get("width", {}).get("magnitude", 612)
        left_margin_pt = doc_style.get("marginLeft", {}).get("magnitude", 72)
        right_margin_pt = doc_style.get("marginRight", {}).get("magnitude", 72)

        # Calculate content width (subtracting margins)
        content_width_pt = page_width_pt - left_margin_pt - right_margin_pt

        # Convert to pixels (1pt = 96/72 = 1.333)
        content_width_px = content_width_pt * (96 / 72)

        # Wrap in HTML
        html_wrapped = f"""
            <div style="
                        max-width: {content_width_px}px; 
                        margin: auto; 
                        padding-left: {left_margin_pt * (96 / 72)}px; 
                        padding-right: {right_margin_pt * (96 / 72)}px;
                        border: 1px solid #000;
            ">
                {html_output}
            </div>
        """
        with open("output.html", "w", encoding="utf-8") as html_file:
            html_file.write(html_wrapped)

        print("✅ HTML content saved to output.html")

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
