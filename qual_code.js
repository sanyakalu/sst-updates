window.QUAL_PYTHON = `
import os, re, io
from pathlib import Path
from docx import Document
import pandas as pd

month = os.environ["SST_MONTH"]
year  = os.environ["SST_YEAR"]
last_sst_path = Path(os.environ["PREVIOUS_SST_PATH"])

def last_sst_doc_headers(file_path):
    doc = Document(file_path)
    return [p.text for p in doc.paragraphs if p.style.name.startswith('Heading')]

all_doc_headers = last_sst_doc_headers(last_sst_path)
index = all_doc_headers.index("Appendices")
all_last_sst_products = all_doc_headers[index + 1:]

windows_products_naming_convention = [x for x in all_last_sst_products if 'Windows' in x]
windows_products = []
for product in windows_products_naming_convention:
    product = re.sub(r"\\b(iot)\\b", "", product, flags=re.IGNORECASE)
    product = re.sub(r"\\b(ltsc)\\b", "LTSB", product, flags=re.IGNORECASE)
    product = re.sub(r"\\bv(\\d+(?:H\\d)?)\\b", r"(version \\1)", product, flags=re.IGNORECASE)
    product = re.sub(r"\\s+", " ", product).strip()
    windows_products.append(product)

def get_release_versions(doc):
    releases = []
    for table in doc.tables:
        release = None
        wp = None
        found = False
        for cell in table.rows[0].cells:
            if found:
                release = repr(cell.text)
                found = False
            if cell.text.strip().lower() == "release:":
                found = True
            if release is not None and cell.text.startswith("PIC iX Servers and Client Application:"):
                wp = cell.text.split("PIC iX Servers and Client Application:")[1].strip()
        if release is not None and wp is not None:
            releases.append([release, wp])
    return releases

def get_table(doc, title_text):
    title_text = title_text.strip().lower()
    for table in doc.tables:
        if len(table.rows) < 2:
            continue
        first_row = [cell.text.strip().lower() for cell in table.rows[0].cells]
        if len(set(first_row)) == 1 and title_text in first_row:
            return table
    raise ValueError(f"Table not found: {title_text}")

doc = Document(last_sst_path)
release_versions = get_release_versions(doc)
cleaned_release_versions = []
for a, b in release_versions:
    b = re.sub(r'\\b(?:enterprise|ltsb|ltsc|iot)\\b', '', b, flags=re.IGNORECASE)
    b = ' '.join(b.split())
    cleaned_release_versions.append((a, b))

def clean(text):
    text = text.replace("\\u200b", "").replace("\\u00a0", " ")
    text = text.replace("\\n\\n", ", ").replace("\\n", " ")
    return " ".join(text.split())

all_sst_updates = []
for table_name in windows_products:
    try:
        table = get_table(doc, table_name)
    except ValueError:
        continue
    post_vuln_row = False
    for row in table.rows:
        if post_vuln_row:
            if row.cells[1].text == "":
                continue
            cell0 = row.cells[0].text.strip()
            if cell0 == "7-Zip":
                m = re.search(r'\\d+(?:\\.\\d+)+', row.cells[4].text)
                kb_or_build = m.group(0) if m else ""
            elif cell0 == "Microsoft Edge":
                matches = re.findall(r'\\((.*?)\\)', row.cells[3].text)
                kb_or_build = matches[0] if matches else ""
            else:
                kb_or_build = row.cells[0].text
            all_sst_updates.append([
                clean(row.cells[1].text),
                clean(kb_or_build).replace("*","").replace("Build","").replace("build", "").strip(),
                next((a for a, b in release_versions if b == table_name), None).strip('"').strip("'"),
                f"PIC iX Servers and Client Application: {table_name}",
                clean(row.cells[5].text),
                clean(row.cells[3].text),
                clean(row.cells[4].text),
            ])
        if 'Vulnerability / Patch ID' in row.cells[0].text:
            post_vuln_row = True

df_qual = pd.DataFrame(all_sst_updates, columns=[
    "Qualification Date",
    "Qualified KB/version or component",
    "PIC iX release",
    "Qualified system",
    "Product",
    "Recommended Customer Action",
    "Notes/Instructions",
])
_buf = io.StringIO()
df_qual.to_csv(_buf, index=False)
qual_registry_csv_output = _buf.getvalue()
`;
