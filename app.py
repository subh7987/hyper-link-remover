import streamlit as st
import pandas as pd
import io
import zipfile
import re
from email import policy
from email.parser import BytesParser
from email.generator import BytesGenerator
from bs4 import BeautifulSoup, NavigableString, Comment

# ------------------- Your Original Logic -------------------
def break_links_in_html_safe(html):
    links_found = False

    vml_pattern = re.compile(r'(<v:roundrect\b[^>]*\bhref=")([^"]+)(")', re.IGNORECASE)
    html, vml_count = vml_pattern.subn(r'\1#\3', html)
    if vml_count > 0: links_found = True

    vshape_pattern = re.compile(r'(<v:shape\b[^>]*\bhref=")([^"]+)(")', re.IGNORECASE)
    html, vshape_count = vshape_pattern.subn(r'\1#\3', html)
    if vshape_count > 0: links_found = True

    a_pattern = re.compile(r'(<a\b[^>]*\bhref=")([^"]+)(")', re.IGNORECASE)
    html, a_count = a_pattern.subn(r'\1#\3', html)
    if a_count > 0: links_found = True

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
    msg = BytesParser(policy=policy.default).parsebytes(file_bytes)

    modified = False
    links_present = False
    html_found = False

    for part in msg.walk():
        if part.get_content_type() == "text/html":
            html_found = True
            html = part.get_content()
            cleaned_html, found_links = break_links_in_html_safe(html)

            if found_links:
                links_present = True
            if cleaned_html != html:
                charset = part.get_content_charset() or "utf-8"
                part.set_content(cleaned_html, subtype="html", charset=charset)
                modified = True

    output_bytes = io.BytesIO()
    BytesGenerator(output_bytes, policy=policy.default.clone(linesep="\r\n")).flatten(msg)
    output_bytes.seek(0)

    if modified:
        reason = "Hyperlinks removed/disabled"
    elif not html_found:
        reason = "No HTML content"
    elif html_found and not links_present:
        reason = "No hyperlink found"
    else:
        reason = "HTML present but unchanged"

    return output_bytes, modified, reason

# ------------------- Streamlit UI -------------------
st.set_page_config(page_title="EML Hyperlink Remover", layout="centered")

st.title("📧 EML Hyperlink Remover")
st.write("Upload `.eml` files to remove or disable hyperlinks while keeping buttons/images intact.")

uploaded_files = st.file_uploader("Upload EML files", type=["eml"], accept_multiple_files=True)

if uploaded_files:
    report_data = []
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for uploaded_file in uploaded_files:
            st.write(f"🔄 Processing: {uploaded_file.name}")
            cleaned_bytes, modified, reason = process_eml_file(uploaded_file.read(), uploaded_file.name)

            zipf.writestr(uploaded_file.name, cleaned_bytes.read())
            report_data.append({
                "Filename": uploaded_file.name,
                "Hyperlinks Removed/Disabled": "Yes" if modified else "No",
                "Reason": reason
            })
            st.success(f"{'✅' if modified else '➖'} {uploaded_file.name} → {reason}")

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
        file_name="hyperlink_removal_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
