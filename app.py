import streamlit as st
import pandas as pd
import io
import zipfile
import re
from email import policy
from email.parser import BytesParser
from email.generator import BytesGenerator
from bs4 import BeautifulSoup, NavigableString, Comment

# ------------------- Helper Functions -------------------

def extract_delivered_to_email(raw_msg):
    # Delivered-To
    match = re.search(r"Delivered-To:\s*<?([\w\.-]+@[\w\.-]+\.\w+)>?", raw_msg, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    # Return-Path
    match = re.search(r"Return-Path:\s*<?([\w\.-]+@[\w\.-]+\.\w+)>?", raw_msg, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    # To
    match = re.search(r"^To:\s*.*<?([\w\.-]+@[\w\.-]+\.\w+)>?", raw_msg, re.IGNORECASE | re.MULTILINE)
    if match:
        return match.group(1).strip()

    return None

def replace_email_everywhere(data, email_addr):
    if not email_addr:
        return data
    patterns = [
        (re.escape(f"<{email_addr}>"), "<[email]>"),
        (re.escape(email_addr), "[email]"),
    ]
    for pattern, replacement in patterns:
        data = re.sub(pattern, replacement, data, flags=re.IGNORECASE)
    return data

def break_links_in_html_safe(html):
    links_found = False

    # Replace href inside <v:roundrect> and <v:shape>
    for pattern in [
        re.compile(r'(<v:roundrect\b[^>]*\bhref=")([^"]+)(")', re.IGNORECASE),
        re.compile(r'(<v:shape\b[^>]*\bhref=")([^"]+)(")', re.IGNORECASE),
        re.compile(r'(<a\b[^>]*\bhref=")([^"]+)(")', re.IGNORECASE),
    ]:
        html, count = pattern.subn(r'\1#\3', html)
        if count > 0:
            links_found = True

    soup = BeautifulSoup(html, "lxml")

    for a in soup.find_all("a"):
        if a.find("img") or "v:roundrect" in str(a).lower() or "v:shape" in str(a).lower():
            if a.has_attr("href") and a["href"] != "#":
                a["href"] = "#"
                links_found = True
        else:
            a.unwrap()
            links_found = True

    url_pattern = re.compile(r'https?://[^\s"<>()]+', re.IGNORECASE)
    for text_node in soup.find_all(string=url_pattern):
        if isinstance(text_node, Comment):
            continue
        new_text = url_pattern.sub("[link removed]", text_node)
        if new_text != text_node:
            text_node.replace_with(NavigableString(new_text))
            links_found = True

    return str(soup), links_found

def process_eml_file(file_bytes, filename):
    raw_text = file_bytes.decode("utf-8", errors="ignore")
    delivered_email = extract_delivered_to_email(raw_text)

    if delivered_email:
        raw_text = replace_email_everywhere(raw_text, delivered_email)

    msg = BytesParser(policy=policy.default).parsebytes(raw_text.encode("utf-8", errors="ignore"))

    modified_links = False
    modified_email = bool(delivered_email)
    html_found = False

    for part in msg.walk():
        ctype = part.get_content_type()
        charset = part.get_content_charset() or "utf-8"

        if ctype == "text/html":
            html_found = True
            html = part.get_content()
            html = replace_email_everywhere(html, delivered_email)
            cleaned_html, found_links = break_links_in_html_safe(html)
            if found_links:
                modified_links = True
            part.set_content(cleaned_html, subtype="html", charset=charset)

        elif ctype == "text/plain":
            text = part.get_content()
            text = replace_email_everywhere(text, delivered_email)
            part.set_content(text, subtype="plain", charset=charset)

    output_bytes = io.BytesIO()
    BytesGenerator(output_bytes, policy=policy.default.clone(linesep="\r\n")).flatten(msg)
    output_bytes.seek(0)

    if not html_found:
        reason = "No HTML content"
    elif modified_links and modified_email:
        reason = "Links removed + Emails masked"
    elif modified_links:
        reason = "Links removed"
    elif modified_email:
        reason = "Emails masked"
    else:
        reason = "No changes"

    return output_bytes, reason

# ------------------- Streamlit UI -------------------

st.set_page_config(page_title="EML Cleaner", layout="centered")
st.title("📧 EML Hyperlink Remover + Email Masker")
st.write("Upload `.eml` files to clean hyperlinks and mask email addresses in one go.")

uploaded_files = st.file_uploader("Upload EML files", type=["eml"], accept_multiple_files=True)

if uploaded_files:
    report_data = []
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for uploaded_file in uploaded_files:
            st.write(f"🔄 Processing: {uploaded_file.name}")
            cleaned_bytes, reason = process_eml_file(uploaded_file.read(), uploaded_file.name)

            zipf.writestr(uploaded_file.name, cleaned_bytes.read())
            report_data.append({
                "Filename": uploaded_file.name,
                "Reason": reason
            })
            st.success(f"✅ {uploaded_file.name} → {reason}")

    df = pd.DataFrame(report_data)

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    excel_buffer.seek(0)

    st.download_button(
        label="⬇ Download Cleaned EMLs (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="cleaned_eml_files.zip",
        mime="application/zip"
    )

    st.download_button(
        label="⬇ Download Processing Report (Excel)",
        data=excel_buffer.getvalue(),
        file_name="processing_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
