// AUTO-GENERATED from ../SSTUpdates-git-workflow.py - do not edit by hand.
window.SST_PYTHON = String.raw`
#package imports
import os
from pathlib import Path
from docx import Document
from pathlib import Path
from shutil import copy2
from docx.enum.section import WD_ORIENT
from datetime import datetime
from dateutil.relativedelta import relativedelta
from copy import deepcopy
import pandas as pd
import ast
import html
import re
from docx.table import _Row


#things to change per user
month = os.environ["SST_MONTH"]
year = os.environ["SST_YEAR"]
last_sst_path = Path(os.environ["PREVIOUS_SST_PATH"])
file_txt = Path(os.environ["UPDATE_FILE_PATH"])

# output stuff
output_directory = Path(os.environ.get("OUTPUT_DIRECTORY") or "output")
output_directory.mkdir(parents=True, exist_ok=True)
output_sst_path = (output_directory/ f"PIC iX_Security_Status_Table_{month}_{year}.docx")
copy2(last_sst_path, output_sst_path)

#global variaables
section_title = r'Description of Product(s) Under Test'
additional_product_list = ['Visual Studio', 'SQL Server', 'Edge Browser', '7-Zip', '.NET']
special_products_general = { 'Windows Malicious Software Removal Tool': 'Microsoft Windows Malicious Software Removal Tool (MSRT)', 
                    'Microsoft Edge': 'Microsoft Edge', 
                    'Microsoft Defender Antivirus antimalware platform': 'Microsoft Defender Antivirus (platform update)',
                    '7-Zip': '7-Zip'}
special_products_versioning_etc = {
    r'SQL Server\s+(\d{4})': r'Microsoft SQL Server \1',
    r'^.*?(\.NET\s+\d+\.\d+).*$': r'\1'}
month_map = {
    "Jan": "January",
    "Feb": "February",
    "Mar": "March",
    "Apr": "April",
    "May": "May",
    "Jun": "June",
    "Jul": "July",
    "Aug": "August",
    "Sep": "September",
    "Oct": "October",
    "Nov": "November",
    "Dec": "December"
}

#INPUT FILE CHECKING
# Input file can be .txt or .docx.
# If it's a .docx, convert it to a .txt file and return the txt path.
def get_txt_file(file_path):
    file_path = Path(file_path)
    ext = file_path.suffix.lower()

    if ext == ".txt":
        return str(file_path)

    elif ext == ".docx":
        doc = Document(file_path)

        txt_path = file_path.with_suffix(".txt")

        with open(txt_path, "w", encoding="utf-8") as f:
            # Write paragraphs
            for p in doc.paragraphs:
                if p.text.strip():
                    f.write(p.text + "\n")

            # Write tables
            for table in doc.tables:
                for row in table.rows:
                    line = "\t".join(cell.text.strip() for cell in row.cells)
                    f.write(line + "\n")

        return str(txt_path)

    else:
        raise ValueError(f"Unsupported file type: {ext}")

file_txt = get_txt_file(Path(file_txt))

#CHECK IF right input file is being used and if the month/year in the input file matches the month/year in the output document
#make sure using right input txt file and sst uppdate doxc file
def validate_input_and_document(input_txt_path, output_sst_path, month, year):
    # Convert configured month/year into a datetime.
    target_date = datetime.strptime(f"{month} {year}","%b %Y",)

    expected_input_date = target_date.strftime("%Y-%m")
    expected_document_date = (target_date - relativedelta(months=1)).strftime("%B %Y")

    #get yyy-mm date from input txt file and validate that it matches the expected date
    with open(input_txt_path, "r", encoding="utf-8") as file:
        input_text = file.read()

    input_dates = set(re.findall(r"(?<!\d)(\d{4}-(?:0[1-9]|1[0-2]))(?!\d)", input_text))

    if not input_dates:
        raise ValueError("No YYYY-MM date values were found in the input file.")

    if input_dates != {expected_input_date}:
        raise ValueError(
            "Input file date validation failed.\n"
            f"Expected: {expected_input_date}\n"
            f"Found: {', '.join(sorted(input_dates))}"
            "Please check that input file is correct and that the month/year in the input file matches the month/year in the output document.")

    #check if existing document is using the correct month and year in the header
    doc = Document(output_sst_path)

    document_dates = set()

    for section in doc.sections:
        header = section.header

        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    matches = re.findall(
                        r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b", cell.text, re.IGNORECASE)

                    document_dates.update(match.title() for match in matches)

    if not document_dates:
        raise ValueError(
            "No 'Month YYYY' date was found in the document headers.")

    if expected_document_date not in document_dates:
        raise ValueError(
            "Document month validation failed.\n"
            f"Expected the existing document to be: "
            f"{expected_document_date}\n"
            f"Found: "
            f"{', '.join(sorted(document_dates))}"
        )

    return doc

#pull past SST doc headers
def last_sst_doc_headers(file_path):
    doc = Document(file_path)

    headers = []

    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            headers.append(para.text)
    
    return headers

all_doc_headers = last_sst_doc_headers(last_sst_path)


index = all_doc_headers.index("Appendices")
all_last_sst_windows_products = all_doc_headers[index + 1:]
print("All found Windows products from the last SST document:")
print(all_last_sst_windows_products) 


#Getting microsoft safe product naming conventions
windows_products_naming_convention = [x for x in all_last_sst_windows_products if 'Windows' in x] #only getting Windows relevant products
windows_products_naming_convention

windows_products = []
for product in windows_products_naming_convention:
    product = re.sub(r"\b(iot)\b", "", product, flags=re.IGNORECASE)
    product = re.sub(r"\b(ltsc)\b", "LTSB", product, flags = re.IGNORECASE)
    product = re.sub(r"\bv(\d+(?:H\d)?)\b", r"(version \1)", product, flags=re.IGNORECASE)
    product = re.sub(r"\s+", " ", product)
    product = product.strip()

    windows_products.append(product)

# input file parsing
pattern = re.compile("|".join(map(re.escape, windows_products)))

rows = []
current_item = None
current_sst_name = None
current_section = None  

with open(file_txt, "r", encoding="utf-8") as f:
    for line in f:
        line = line.strip()

        match = pattern.search(line)

        if match and match.start() < 10:
            current_item = match.group()

            # Capture text before product name (e.g. "c.03", "4.7.2")
            current_section = line[:match.start()].strip()

            # Find corresponding naming convention
            idx = windows_products.index(current_item)
            current_sst_name = windows_products_naming_convention[idx]

        elif current_item and line:
            rows.append([
                current_section,   
                current_sst_name,
                current_item,
                line
            ])

df = pd.DataFrame(rows,columns=["Section","SST Naming Conv","Product","KB Updates"])

#get KB numbers and CVE numbers from the 'KB Updates' column, remove duplicates, and store them in a new column 'KB numbers' + remove htmls if there are any n the text
df['KB numbers'] = df['KB Updates'].apply(lambda x: list(dict.fromkeys( re.findall(r'KB\d+|CVE-\d{4}-\d+', str(x))))) #extract CVE numbers and KB numbers from the update list
df['KB Updates'] = df['KB Updates'].str.replace(r'https?://\S+|www\.\S+', '', regex=True) # drop URLS from text

#get KB numbers and CVE numbers from the 'KB Updates' column, remove duplicates, and store rest of the remaining information in Notes/Instrunctions column
df['Notes / Instrunctions'] = df['KB Updates'].apply(lambda x: re.sub(r'KB\d+|CVE-\d{4}-\d+', '', x).strip())
df['Notes / Instrunctions'] = df['Notes / Instrunctions'].apply(lambda x: re.sub(r"\b\w{4}-\w{2}\s+", "", x))

# Move installer/build rows into the nearest non-move row above - so if there's no KB or CVE found - assume likely some type of install/package instrunctions - find next available row above that has a KB or CVE and move the text there. If there's no row above with a KB or CVE, then leave it as is.
move_mask = df["KB Updates"].str.contains(
    r"\.exe\b|\(Build.*?\)",
    case=False,
    regex=True,
    na=False
)

df["Recommended Customer Action"] = ""

df["Recommended Customer Action"] = df["Recommended Customer Action"].fillna("")

rows_to_drop = []

for idx in df[move_mask].index:

    target = idx - 1

    # Find nearest row above that is NOT a move row
    while target >= 0 and move_mask.iloc[target]:
        target -= 1

    if target < 0:
        continue

    value = str(df.loc[idx, "KB Updates"]).strip()
    
    if '.exe' in value.lower(): #check if there's any leading values before install packages (like a.) or a) etc)
        match = re.search(r"([^\s]+\.exe)\b", value, flags=re.IGNORECASE)
        if match:
            value = match.group(1)


    text = f"Install: {value}"

    # Append cleanly
    if df.loc[target, "Recommended Customer Action"]:
        df.loc[target, "Recommended Customer Action"] += "\n\n" + text
    else:
        df.loc[target, "Recommended Customer Action"] = text

    rows_to_drop.append(idx)

df = df.drop(rows_to_drop).reset_index(drop=True)

#now drop any rows that have no KB or CVE numbers and no Notes/Instructions - these are essentially empty rows
rows_to_drop = []
rows_to_drop.extend(df.index[df['KB Updates'].fillna('').astype(str).str.strip().eq('')].tolist()) #dropping empty lists essentially

df = df.drop(index=list(set(rows_to_drop))).reset_index(drop=True)


#formatting errors - if leading spaces, periods, dashes, remove
df['Notes / Instrunctions'] = df['Notes / Instrunctions'].str.replace(r'^[\s\-\•]+','',regex=True)

#clean up - if there's empty parentheses in the Notes/Instructions column, remove them
df["Notes / Instrunctions"] = df["Notes / Instrunctions"].str.replace("()", "", regex=False)

#same thing as before but looking specifically for rows that have no KB or CVE numbers and are not known install things (.exe or Build things)
# Find candidate rows
empty_rows = df[
    df["KB numbers"].apply(lambda x: isinstance(x, list) and len(x) == 0)
    & (
        df["Recommended Customer Action"].isna()
        | df["Recommended Customer Action"].eq("")
    )
]

rows_to_drop = []

for idx, row in empty_rows.iterrows():

    notes = str(row["Notes / Instrunctions"])
    product = row["Product"]

    # Find which additional product appears in the notes
    matching_products = [
        p for p in additional_product_list
        if p in notes
    ]

    if not matching_products:
        continue

    for add_product in matching_products:

        # Find target rows:
        # same Product
        # has KB numbers
        # notes also contain the same additional product
        target_mask = (
            (df["Product"] == product)
            & df["KB numbers"].apply(
                lambda x: isinstance(x, list) and len(x) > 0
            )
            & df["Notes / Instrunctions"].fillna("").str.contains(
                add_product, regex=False
            )
        )

        target_rows = df[target_mask]

        if target_rows.empty:
            continue

        target_idx = target_rows.index[0]

        # Append Notes / Instructions to Recommended Customer Action
        existing_action = df.at[target_idx, "Recommended Customer Action"]

        if pd.isna(existing_action):
            existing_action = ""

        if existing_action:
            df.at[target_idx, "Recommended Customer Action"] = (
                f"Install: {existing_action}\n\nInstall: {notes}"
            )
        else:
            df.at[target_idx, "Recommended Customer Action"] = notes

        rows_to_drop.append(idx)

        # Stop after first successful match
        break

# Delete merged rows
df = df.drop(index=set(rows_to_drop)).reset_index(drop=True)

#convert KB numbers list to string for easier processing in the next steps
df['KB numbers'] = df['KB numbers'].apply(lambda x: x[0] if isinstance(x, list) and len(x) > 0 else '') 

#check if notes/instrunctions starts with asterik, if so move to KB numbers column and remove from Notes/Instructions column
mask = df["Notes / Instrunctions"].str.match(r"^\*+", na=False)

# Extract leading asterisks
stars = df.loc[mask, "Notes / Instrunctions"].str.extract(r"^(\*+)", expand=False)

# Remove them from col1
df.loc[mask, "Notes / Instrunctions"] = df.loc[mask, "Notes / Instrunctions"].str.replace(r"^\*+", "", regex=True)

# Append them to col2
df.loc[mask, "KB numbers"] = (df.loc[mask, "KB numbers"].fillna("").astype(str) + stars)

#for the Notes / Instructions column, if it contains any of the following phrases, replace the entire value with that phrase (to shorten the text)
notes_instrunctions_shortening = ["Update for Microsoft Defender Antivirus antimalware platform", "Servicing Stack Update for Windows", "Cumulative Update for Windows"]
pattern = "|".join(re.escape(x) for x in notes_instrunctions_shortening)
df['Notes / Instrunctions'] = df['Notes / Instrunctions'].str.extract(f"({pattern})", expand=False).fillna(df["Notes / Instrunctions"]) #extract the first matching pattern from the list, if none found, keep the original value

# get special exceptions for microsoft edge
mask = df['Notes / Instrunctions'].eq("Edge Browser Security Update")
build = df.loc[mask, 'Recommended Customer Action'].str.extract(r'(\(Build[^)]*\))', expand=False)
df.loc[mask, 'Notes / Instrunctions'] = "Microsoft Edge " + build

#make recommended customer default 'Install Update' if no specific instrunctions found
df['Recommended Customer Action'] = df['Recommended Customer Action'].replace("", "Install Recommended Update").fillna("Install Recommended Update")

#make sure service stack update is always listed first
df["_ssu_last"] = df["Notes / Instrunctions"].str.contains("Servicing Stack Update", case=False, na=False)

df = df.sort_values(by=["Product", "_ssu_last"], ascending=[True, True]).drop(columns="_ssu_last")

#get microsoft products column to match the SST namingproductstring convention for Microsoft products
def simplifying_windows_products_for_sst_col(productstring):
    product = re.sub(r"\b(enterprise|ltsb|ltsc)\b", "", productstring, flags=re.IGNORECASE)
    product = re.sub(r"\s+", " ", product)
    product = re.sub(r"\(([^)]*?)\bversion\b([^)]*?)\)",r"\1Version\2",product,flags=re.IGNORECASE)
    product = product.strip()
    return product

df['Windows Product'] = df['Product'].apply(lambda x: simplifying_windows_products_for_sst_col(x) if isinstance(x, str) else x)

#if kb appears across multiple products, make sure all are reflected in the 'Windows Product' column for that KB number in every table
kb_products = (df.groupby("KB numbers")["Windows Product"].transform(lambda x: "\n\n".join(sorted(set(x.dropna())))))

df["Windows Product"] = kb_products

#for speciality products (products outside header items - fill in microsoft product category)
#this special products is specifically for products in which the versioning or specific build or year  is of relevance to be added to the microsoft product
for pattern, replacement in special_products_versioning_etc.items():
    mask = df["Notes / Instrunctions"].str.contains(pattern, regex=True, na=False)
    df.loc[mask, "Windows Product"] = (df.loc[mask, "Notes / Instrunctions"].str.replace(pattern, replacement, regex=True))

#for speciality naming systems where versioning year not relevant
for pattern, replacement in special_products_general.items():
    mask = df["Notes / Instrunctions"].str.contains(pattern, regex=True, na=False)
    df.loc[mask, "Windows Product"] = replacement
    #check if kb number is empty, then if so fill
    empty_product = df['KB numbers'].fillna("").str.strip().eq("")
    df.loc[mask & empty_product, 'KB numbers'] = replacement


def get_table(doc, title_text):
    title_text = title_text.strip().lower()

    for table in doc.tables:
        if len(table.rows) < 2:
            continue

        first_row = [cell.text.strip().lower() for cell in table.rows[0].cells]

        if len(set(first_row)) == 1 and title_text in first_row:
            return table

    raise ValueError(f"Target table not found: {title_text}")

def insert_rows_with_format(table, new_rows):
    tbl = table._element

    header_tr = table.rows[1]._tr
    header_index = list(tbl).index(header_tr)

    template_tr = table.rows[2]._tr

    for row_data in new_rows:
        insert_index = header_index + 1

        new_tr = deepcopy(template_tr)
        tbl.insert(insert_index, new_tr)

        new_row = _Row(new_tr, table)

        for j, val in enumerate(row_data):
            if j < len(new_row.cells):

                if val is None:
                    text = ""
                elif isinstance(val, float) and pd.isna(val):
                    text = ""
                else:
                    text = str(val)

                new_row.cells[j].text = text


doc = Document(last_sst_path)

for table_name, group in df.groupby("SST Naming Conv"):
    print(f"Processing table: {table_name}")
    try:
        table = get_table(doc, table_name)
    except ValueError:
        print(f"Table not found: {table_name}")
        continue

    new_rows = []

    for _, row in group.iterrows():
        new_rows.append([
            row["KB numbers"],
            "",
            "Solution Available",
            row['Recommended Customer Action'],
            row["Notes / Instrunctions"],
            row["Windows Product"]
        ])

    insert_rows_with_format(table, new_rows)

    kb_num = None
    found_first_update = False
    # Replace the newest SSU row
    for row in table.rows:
        if "Servicing Stack Update for Windows" in (row.cells[4].text):
            if not found_first_update:
                kb_num = row.cells[0].text  # KB is in first column
                found_first_update = True
                continue
            if found_first_update:
                row.cells[3].text = f"Superseded by {kb_num}"
                break


doc.save(output_sst_path)

#changing document header to reflect the new month and year

doc = Document(output_sst_path)
doc = validate_input_and_document(input_txt_path=file_txt,output_sst_path=output_sst_path,month=month,year=year)
newest_reving = None


# Set every document section to landscape
for section in doc.sections:
    section.orientation = WD_ORIENT.LANDSCAPE

    # if dimensions still portrait (outlirs) - make sure swaps dimensions to landscape mode
    if section.page_height > section.page_width:
        section.page_width, section.page_height = (
            section.page_height,
            section.page_width,
        )

#get next doc revision (ex: Dp -> Dq)
def next_revision(rev):
    chars = list(rev.upper())

    i = len(chars) - 1
    while i >= 0:
        if chars[i] < "Z":
            chars[i] = chr(ord(chars[i]) + 1)
            break

        chars[i] = "A"
        i -= 1

    if i < 0:
        chars.insert(0, "A")

    return "".join(chars)


# Update headers in every section
processed_headers = set()

for section in doc.sections:
    for header in [section.header, section.first_page_header, section.even_page_header]:

        header_id = id(header.part)

        # Prevent linked headers from being updated more than once
        if header_id in processed_headers:
            continue

        processed_headers.add(header_id)

        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text

                    # Revision
                    match = re.search(
                        r"\bRev(?:ision)?\.?\s*([A-Z]{2,})\b",text,re.IGNORECASE)

                    if match:
                        old_rev = match.group(1)
                        new_rev = next_revision(old_rev)
                        newest_reving = new_rev

                        text = (
                            text[:match.start(1)] + new_rev + text[match.end(1):])

                    # Month and year
                    text = re.sub(r"[A-Z][a-z]+\s+\d{4}", f"{month_map[month]} {year}", text)

                    cell.text = text


# Update the revision table in the document body
for table in doc.tables:
    for row_index, row in enumerate(table.rows):
        if len(row.cells) >= 2:
            if (
                row.cells[0].text.strip() == "Revision"
                and row.cells[1].text.strip()
                == "Change Description in Brief"
            ):
                # Make sure there is a row after the header row
                if row_index + 1 < len(table.rows):
                    revision_row = table.rows[row_index + 1]

                    if newest_reving is not None:
                        revision_row.cells[0].text = newest_reving

                    revision_row.cells[1].text = (
                        f"Update for {month_map[month]} "
                        "patch qualifications."
                    )

                break


doc.save(output_sst_path)



`;
