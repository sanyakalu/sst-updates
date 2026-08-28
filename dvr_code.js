// AUTO-GENERATED from dvr_updates.py — do not edit by hand.
window.DVR_PYTHON = String.raw`
import calendar
import re
import html as html_module
from datetime import datetime
import pandas as pd
import numpy as np
import json as _json
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

filter_month = int(os.environ["DVR_MONTH_NUM"])
filter_year  = int(os.environ["DVR_YEAR"])

output_directory = Path(os.environ.get("OUTPUT_DIRECTORY", "/work/output"))
output_directory.mkdir(parents=True, exist_ok=True)

_test_rows       = _json.loads(dvr_test_rows_json)
_sticr_items_raw = _json.loads(dvr_sticr_items_json)

doc = Document("/work/dvr_template.docx")

# ── Fill {MONTH}/{YEAR} placeholders in all tables ────────────────────────────
for _tbl in doc.tables:
    for _row in _tbl.rows:
        for _cell in _row.cells:
            if "{MONTH}" in _cell.text or "{YEAR}" in _cell.text:
                _cell.text = (
                    _cell.text
                    .replace("{MONTH}", calendar.month_name[filter_month])
                    .replace("{YEAR}", str(filter_year))
                )

# ── Update section header title ───────────────────────────────────────────────
_new_title = (
    f"PIC iX Security Patch Design Verification Report "
    f"{calendar.month_name[filter_month]} {filter_year} Update"
)
_processed_hdrs = set()
for _sec in doc.sections:
    if id(_sec.header.part) in _processed_hdrs:
        continue
    _processed_hdrs.add(id(_sec.header.part))
    for _tbl in _sec.header.tables:
        for _row in _tbl.rows:
            for _cell in _row.cells:
                for _para in _cell.paragraphs:
                    if re.search(r'Design Verification Report', _para.text, re.IGNORECASE):
                        for _i, _run in enumerate(_para.runs):
                            _run.text = _new_title if _i == 0 else ""

# ── Constants ─────────────────────────────────────────────────────────────────
SECTIONS = {
    "Install Ancillary Software Update": ["install ancillary software"],
    "Performance Test":                  ["longevity", "performance"],
    "Functional Test":                   ["functional", "microsoft edge", "antivirus", "malware"],
    "Reliability Test":                  ["reliability", "reboot"],
}
SECTION_ORDER = [
    "Install Ancillary Software Update",
    "Performance Test",
    "Functional Test",
    "Reliability Test",
]
PRODUCT_PAIRS = [
    ['1607', '2016', 'Windows 10 1607, Windows Server 2016',  'Fixture intF'],
    ['1809', '2019', 'Windows 10 1809, Windows Server 2019',  'Fixture ex99'],
    ['21H2', '2022', 'Windows 10 21H2, Windows Server 2022',  'Fixture ex83'],
]

# ── Build df from injected test rows ──────────────────────────────────────────
_cols = ['Plan ID', 'Plan Name', 'Run Name', 'Run ID', 'Run State',
         'Outcome', 'Test Case Name', 'Test Case ID']
df = pd.DataFrame(_test_rows or [], columns=_cols)

if not df.empty:
    df['Run Name'] = df['Run Name'].str.replace(r'\(Manual\)', '', regex=True).str.strip()
    df['Version']  = df['Run Name'].str.extract(r'([A-Za-z0-9]*\.[A-Za-z0-9.]+)', expand=False)
    df['Doc Table'] = np.where(
        df['Run Name'].str.contains('iX', na=False),
        'Test Summary-OS Security Updates',
        'Test Summary-' + df['Run Name'].fillna(''),
    )
else:
    df['Run Name']  = pd.Series(dtype=str)
    df['Version']   = pd.Series(dtype=str)
    df['Doc Table'] = pd.Series(dtype=str)

def determine_section(test_case_name):
    name = str(test_case_name).lower()
    for section, patterns in SECTIONS.items():
        for pattern in patterns:
            if pattern.lower() in name:
                return section
    return "Other"

df["Section"] = df["Test Case Name"].apply(determine_section)

doc_df = pd.DataFrame({
    "Table":           df.get('Doc Table',      pd.Series(dtype=str)),
    "Section":         df.get('Section',        pd.Series(dtype=str)),
    "Test Case ID":    df.get('Test Case ID',   pd.Series(dtype=str)),
    "Test Case Title": df.get('Test Case Name', pd.Series(dtype=str)),
    "PIC iX Build":    df.get('Version',        pd.Series(dtype=str)),
    "Run Name":        df.get('Run Name',       pd.Series(dtype=str)),
    "Run ID":          df.get('Run ID',         pd.Series(dtype=str)),
    "Result":          df.get('Outcome',        pd.Series(dtype=str)),
})

# ── Find insert point ─────────────────────────────────────────────────────────
_target_section = "Design Verification Results"
insert_after = None
for _para in doc.paragraphs:
    if _para.text.strip() == _target_section:
        insert_after = _para._element
        break
if insert_after is None:
    raise ValueError(f"Section '{_target_section}' not found in document.")

# ── Helper functions ──────────────────────────────────────────────────────────
def set_cell_font(cell, font_name="Calibri", size=10, bold=False):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size)
            run.bold = bold
            try:
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                _W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                rFonts.set(_W + "ascii", font_name)
                rFonts.set(_W + "hAnsi", font_name)
            except AttributeError:
                pass

def shade_cell(cell, color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def get_ancillary_solution(run_name):
    run_lower = str(run_name).lower()
    for pair in PRODUCT_PAIRS:
        a, b, label = pair[0], pair[1], pair[2]
        if a.lower() in run_lower and b.lower() in run_lower:
            return label
    return ""

def get_verification_equipment(product_name):
    m = re.search(r'\(([^)]+)\)', str(product_name))
    if not m:
        return ""
    parens_lower = m.group(1).lower()
    fixtures = []
    for pair in PRODUCT_PAIRS:
        a, b, fixture = pair[0], pair[1], pair[3]
        if a.lower() in parens_lower or b.lower() in parens_lower:
            if fixture not in fixtures:
                fixtures.append(fixture)
    return ", ".join(fixtures)

def create_test_table(doc, table_df):
    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    headers = [
        "Test Case ID", "Test Case Title", "PIC iX Build",
        "Ancillary Solution", "Run ID", "Qualitative Passed/Failed Result",
    ]
    for cell, header in zip(table.rows[0].cells, headers):
        cell.text = header
        shade_cell(cell, "F2F2F2")
        set_cell_font(cell, bold=True)
    for section in SECTION_ORDER:
        section_df = table_df[table_df["Section"] == section]
        if section_df.empty:
            continue
        row = table.add_row()
        merged = row.cells[0]
        for cell in row.cells[1:]:
            merged = merged.merge(cell)
        merged.text = section
        shade_cell(merged, "B8CCE4")
        set_cell_font(merged, bold=True)
        for _, item in section_df.iterrows():
            cells = table.add_row().cells
            cells[0].text = str(item["Test Case ID"])
            cells[1].text = str(item["Test Case Title"])
            cells[2].text = str(item["PIC iX Build"])
            cells[3].text = get_ancillary_solution(item["Run Name"])
            cells[4].text = str(item["Run ID"])
            cells[5].text = str(item["Result"])
            for cell in cells:
                set_cell_font(cell)
    widths = [0.8, 2.4, 1.0, 1.8, 0.9, 1.2]
    for col, width in zip(table.columns, widths):
        for cell in col.cells:
            cell.width = Inches(width)

# ── Plan summary table ────────────────────────────────────────────────────────
plan_summary = df[['Plan ID', 'Plan Name']].drop_duplicates() if not df.empty else pd.DataFrame(columns=['Plan ID', 'Plan Name'])

summary_table = doc.add_table(rows=1, cols=3)
summary_table.style = "Table Grid"
for cell, header in zip(summary_table.rows[0].cells,
                         ["Azure DevOps Test Plans Instance Name", "Test Plan ID", "Title"]):
    cell.text = header
    shade_cell(cell, "F2F2F2")
    set_cell_font(cell, bold=True)
for _, row in plan_summary.iterrows():
    cells = summary_table.add_row().cells
    cells[0].text = "https://dev.azure.com/PhilipsMA"
    cells[1].text = str(row["Plan ID"])
    cells[2].text = str(row["Plan Name"])
    for cell in cells:
        set_cell_font(cell)

insert_after.addnext(summary_table._element)
plan_summary_elem = summary_table._element

# ── Bullet list ───────────────────────────────────────────────────────────────
current = plan_summary_elem
doc_df_copy = doc_df.copy()
doc_df_copy["Ancillary Solution"] = doc_df_copy["Run Name"].apply(get_ancillary_solution)

bullet_data = []
_ancillary_groups = doc_df_copy[doc_df_copy["Ancillary Solution"] != ""].groupby("Ancillary Solution", sort=False)
for ancillary, group in _ancillary_groups:
    build = group["PIC iX Build"].dropna().iloc[0] if not group["PIC iX Build"].dropna().empty else ""
    tables = list(dict.fromkeys(
        t.replace("Test Summary-", "").strip()
        for t in group["Table"].dropna()
    ))
    bullet_data.append((build, ancillary, tables))

pair_order = {pair[2]: i for i, pair in enumerate(PRODUCT_PAIRS)}
bullet_data.sort(key=lambda x: pair_order.get(x[1], 999))

intro = doc.add_paragraph("The associated software revisions are as below:")
for run in intro.runs:
    run.font.size = Pt(8)
intro.paragraph_format.space_after  = Pt(6)
intro.paragraph_format.space_before = Pt(6)
current.addnext(intro._element)
current = intro._element

has_list_bullet  = "List Bullet"   in [s.name for s in doc.styles]
has_list_bullet2 = "List Bullet 2" in [s.name for s in doc.styles]

for build, ancillary, tables in bullet_data:
    if has_list_bullet:
        p1 = doc.add_paragraph(f"PIC iX {build}: {ancillary}", style="List Bullet")
    else:
        p1 = doc.add_paragraph(f"•    PIC iX {build}: {ancillary}")
    for run in p1.runs:
        run.font.size = Pt(8)
    current.addnext(p1._element)
    current = p1._element
    if has_list_bullet2:
        p2 = doc.add_paragraph(f"Applicable to: {', '.join(tables)}", style="List Bullet 2")
    else:
        p2 = doc.add_paragraph(f"o    Applicable to: {', '.join(tables)}")
    for run in p2.runs:
        run.font.size = Pt(8)
    p2.paragraph_format.space_after = Pt(6)
    current.addnext(p2._element)
    current = p2._element

# ── Test summary tables ───────────────────────────────────────────────────────
for table_name in doc_df["Table"].dropna().unique():
    table_df = doc_df[doc_df["Table"] == table_name]
    heading = doc.add_heading(table_name, level=4)
    current.addnext(heading._element)
    current = heading._element
    create_test_table(doc, table_df)
    table_elem = doc.tables[-1]._element
    current.addnext(table_elem)
    current = table_elem
    spacer = doc.add_paragraph()
    current.addnext(spacer._element)
    current = spacer._element

# ── STICR processing from injected data ──────────────────────────────────────
_additional_product_list = ['Visual Studio', 'SQL Server', 'Edge Browser', '7-Zip', '.NET']
_notes_shortening = [
    "Update for Microsoft Defender Antivirus antimalware platform",
    "Servicing Stack Update for Windows",
    "Cumulative Update for Windows",
    "Cumulative Update for .NET Framework",
]
_product_header_re = re.compile(r'^(\d+\.\d+(?:\.\d+)?)\s*-\s*(.+)$')

def html_to_text(html_str):
    text = re.sub(r'<br\s*/?>|</p>|</li>|</div>', '\n', html_str, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = html_module.unescape(text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def extract_updates_section(plain_text, yr, mo):
    _month_names = {
        1:'January', 2:'February', 3:'March', 4:'April',
        5:'May', 6:'June', 7:'July', 8:'August',
        9:'September', 10:'October', 11:'November', 12:'December',
    }
    start_match = re.search(
        rf"{_month_names[mo]}\s+{yr}\s+Security Updates",
        plain_text, re.IGNORECASE
    )
    if not start_match:
        start_match = re.search(r'^\d+\.\d+(?:\.\d+)?\s*-\s*Windows', plain_text, re.MULTILINE)
    if not start_match:
        return plain_text
    section = plain_text[start_match.start():]
    end_match = re.search(r'\b\d{1,2}-[A-Za-z]{3,9}-\d{4}\b', section)
    if end_match:
        section = section[:end_match.start()]
    return section.strip()

_sticr_rows = []
for _si in _sticr_items_raw:
    _sid  = _si["id"]
    _desc = _si.get("description", "") or ""
    _section_text = extract_updates_section(html_to_text(_desc), filter_year, filter_month)
    _cur_section = None
    _cur_product = None
    for _line in _section_text.splitlines():
        _line = _line.strip()
        if not _line or not re.search(r'[a-zA-Z0-9]', _line):
            continue
        _hm = _product_header_re.match(_line)
        if _hm:
            _cur_section = _hm.group(1)
            _cur_product = _hm.group(2).strip()
        elif _cur_product:
            _sticr_rows.append([_cur_section, _cur_product, _line, _sid])

sticr_df = pd.DataFrame(_sticr_rows or [], columns=["Section", "Product", "KB Updates", "STICR ID"])

if not sticr_df.empty:
    sticr_df['KB numbers'] = sticr_df['KB Updates'].apply(
        lambda x: list(dict.fromkeys(re.findall(r'KB\d+|CVE-\d{4}-\d+', str(x))))
    )
    sticr_df['KB Updates'] = sticr_df['KB Updates'].str.replace(
        r'https?://\S+|www\.\S+', '', regex=True
    )
    sticr_df['Notes / Instructions'] = sticr_df['KB Updates'].apply(
        lambda x: re.sub(r'KB\d+|CVE-\d{4}-\d+', '', x).strip()
    )
    sticr_df['Notes / Instructions'] = sticr_df['Notes / Instructions'].apply(
        lambda x: re.sub(r'\b\w{4}-\w{2}\s+', '', x)
    )

    # Fold .exe / Build rows (except Microsoft Edge sub-lines)
    _move_mask = (
        sticr_df["KB Updates"].str.contains(r'\.exe\b|\(Build.*?\)', case=False, regex=True, na=False)
        & ~sticr_df["KB Updates"].str.match(r'^Microsoft Edge', case=False, na=False)
    )
    sticr_df["Recommended Customer Action"] = ""
    _rows_to_drop = []
    for _idx in sticr_df[_move_mask].index:
        _target = _idx - 1
        while _target >= 0 and _move_mask.iloc[_target]:
            _target -= 1
        if _target < 0:
            continue
        _val = str(sticr_df.loc[_idx, "KB Updates"]).strip()
        if '.exe' in _val.lower():
            _m = re.search(r"([^\s]+\.exe)\b", _val, flags=re.IGNORECASE)
            if _m:
                _val = _m.group(1)
        _text = f"Install: {_val}"
        _existing = sticr_df.loc[_target, "Recommended Customer Action"]
        sticr_df.loc[_target, "Recommended Customer Action"] = (
            (_existing + "\n\n" + _text) if _existing else _text
        )
        _rows_to_drop.append(_idx)
    sticr_df = sticr_df.drop(_rows_to_drop).reset_index(drop=True)

    sticr_df = sticr_df.drop(
        index=sticr_df.index[sticr_df['KB Updates'].fillna('').str.strip().eq('')].tolist()
    ).reset_index(drop=True)

    sticr_df['Notes / Instructions'] = (
        sticr_df['Notes / Instructions']
        .str.replace(r'^[\s\-\•]+', '', regex=True)
        .str.replace("()", "", regex=False)
    )

    # Fold empty rows into additional product targets
    _empty_rows = sticr_df[
        sticr_df["KB numbers"].apply(lambda x: isinstance(x, list) and len(x) == 0)
        & (sticr_df["Recommended Customer Action"].isna() | sticr_df["Recommended Customer Action"].eq(""))
        & ~sticr_df["KB Updates"].str.match(r'^Microsoft Edge', case=False, na=False)
    ]
    _rows_to_drop = []
    for _idx, _row in _empty_rows.iterrows():
        _notes = str(_row["Notes / Instructions"])
        for _add_prod in [p for p in _additional_product_list if p in _notes]:
            _target_mask = (
                (sticr_df["Product"] == _row["Product"])
                & sticr_df["KB numbers"].apply(lambda x: isinstance(x, list) and len(x) > 0)
                & sticr_df["Notes / Instructions"].fillna("").str.contains(_add_prod, regex=False)
            )
            _target_rows = sticr_df[_target_mask]
            if _target_rows.empty:
                continue
            _target_idx = _target_rows.index[0]
            _existing = sticr_df.at[_target_idx, "Recommended Customer Action"] or ""
            sticr_df.at[_target_idx, "Recommended Customer Action"] = (
                f"Install: {_existing}\n\nInstall: {_notes}" if _existing else _notes
            )
            _rows_to_drop.append(_idx)
            break
    sticr_df = sticr_df.drop(index=set(_rows_to_drop)).reset_index(drop=True)

    sticr_df['KB numbers'] = sticr_df['KB numbers'].apply(
        lambda x: x[0] if isinstance(x, list) and len(x) > 0 else ''
    )

    _star_mask = sticr_df["Notes / Instructions"].str.match(r"^\*+", na=False)
    _stars = sticr_df.loc[_star_mask, "Notes / Instructions"].str.extract(r"^(\*+)", expand=False)
    sticr_df.loc[_star_mask, "Notes / Instructions"] = (
        sticr_df.loc[_star_mask, "Notes / Instructions"].str.replace(r"^\*+", "", regex=True)
    )
    sticr_df.loc[_star_mask, "KB numbers"] = (
        sticr_df.loc[_star_mask, "KB numbers"].fillna("").astype(str) + _stars
    )

    _short_pattern = "|".join(re.escape(x) for x in _notes_shortening)
    sticr_df['Notes / Instructions'] = (
        sticr_df['Notes / Instructions']
        .str.extract(f"({_short_pattern})", expand=False)
        .fillna(sticr_df["Notes / Instructions"])
    )

    # Drop "Edge Browser Security Update" parent rows
    sticr_df = sticr_df[
        ~sticr_df["Notes / Instructions"].str.contains(
            "Edge Browser Security Update", case=False, na=False
        )
    ].reset_index(drop=True)

    # Parse "Microsoft Edge XYZ Version ..." lines
    def parse_edge_line(line):
        m = re.match(r'^(Microsoft Edge[\w\s\-]*?)\s+(Version\s+.+)$', str(line), re.IGNORECASE)
        if m:
            return m.group(1).strip(), m.group(2).strip()
        return "Microsoft Edge", str(line)

    _edge_mask = sticr_df["KB Updates"].str.match(r'^Microsoft Edge', case=False, na=False)
    if _edge_mask.any():
        _parsed = sticr_df.loc[_edge_mask, "KB Updates"].apply(parse_edge_line)
        sticr_df.loc[_edge_mask, "KB numbers"]          = _parsed.apply(lambda x: x[0])
        sticr_df.loc[_edge_mask, "Notes / Instructions"] = _parsed.apply(lambda x: x[1])

    sticr_df['Recommended Customer Action'] = (
        sticr_df['Recommended Customer Action']
        .replace("", "Install Recommended Update")
        .fillna("Install Recommended Update")
    )

    sticr_df["_ssu"] = sticr_df["Notes / Instructions"].str.contains(
        "Servicing Stack Update", case=False, na=False
    )
    sticr_df = sticr_df.sort_values(
        ["Product", "_ssu"], ascending=[True, True]
    ).drop(columns="_ssu").reset_index(drop=True)

    def shorten_product(p):
        p = re.sub(r'\b(enterprise|ltsb|ltsc|iot|version)\b', '', p, flags=re.IGNORECASE)
        p = re.sub(r'[()]', '', p)
        return re.sub(r'\s+', ' ', p).strip()

    sticr_df["Short Product"] = sticr_df["Product"].apply(shorten_product)
    _kb_products = sticr_df.groupby("KB numbers")["Short Product"].transform(
        lambda x: ", ".join(sorted(set(x.dropna())))
    )
    sticr_df["Product Name"] = sticr_df["KB numbers"] + " (" + _kb_products + ")"
    _no_key_mask = sticr_df["KB numbers"].eq("")
    sticr_df.loc[_no_key_mask, "Product Name"] = sticr_df.loc[_no_key_mask, "Notes / Instructions"]

    sticr_df = sticr_df.drop_duplicates(
        subset=["KB numbers", "Notes / Instructions"], keep="first"
    ).reset_index(drop=True)
else:
    sticr_df = pd.DataFrame(columns=["Section", "Product", "KB Updates", "STICR ID",
                                      "KB numbers", "Notes / Instructions",
                                      "Recommended Customer Action", "Product Name"])

print(f"STICR rows after processing: {len(sticr_df)}")

# ── STICR table ───────────────────────────────────────────────────────────────
def has_version_info(notes):
    return bool(re.search(r'v\d+[.\d]*|\d+\.\d+|Build|\.exe', str(notes), re.IGNORECASE))

def numbered_installs(rca_text):
    items = [item.strip() for item in re.split(r'Install:\s*', str(rca_text)) if item.strip()]
    return "\n".join(f"{i}. {item}" for i, item in enumerate(items, 1))

def create_sticr_table(doc, sticr_df):
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    headers = ["Product Name", "Product Number", "Software Revision",
               "Verification Equipment", "STICR ID"]
    for cell, header in zip(table.rows[0].cells, headers):
        cell.text = header
        shade_cell(cell, "F2F2F2")
        set_cell_font(cell, bold=True)
    for _, item in sticr_df.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(item["Product Name"])
        cells[1].text = "N/A"
        _notes = str(item["Notes / Instructions"])
        _rca   = str(item.get("Recommended Customer Action", ""))
        _has_exe = bool(re.search(r'\.exe', _rca, re.IGNORECASE))
        if has_version_info(_notes):
            cells[2].text = _notes + ("\n" + numbered_installs(_rca) if _has_exe else "")
        elif _has_exe:
            cells[2].text = numbered_installs(_rca)
        else:
            cells[2].text = "N/A"
        cells[3].text = get_verification_equipment(str(item["Product Name"]))
        cells[4].text = str(item["STICR ID"])
        for cell in cells:
            set_cell_font(cell)
    widths = [2.5, 1.0, 2.5, 1.2, 0.8]
    for col, width in zip(table.columns, widths):
        for cell in col.cells:
            cell.width = Inches(width)

_sticr_heading = doc.add_heading("Description of Product(s) Under Test", level=3)
current.addnext(_sticr_heading._element)
current = _sticr_heading._element

create_sticr_table(doc, sticr_df)
_sticr_table_elem = doc.tables[-1]._element
current.addnext(_sticr_table_elem)
current = _sticr_table_elem

_spacer = doc.add_paragraph()
current.addnext(_spacer._element)

# ── Save ──────────────────────────────────────────────────────────────────────
dvr_output_filename = f"DVR_PIC_iX_{calendar.month_name[filter_month]}_{filter_year}.docx"
doc.save(str(output_directory / dvr_output_filename))
print(f"Saved: {dvr_output_filename}")
`;
