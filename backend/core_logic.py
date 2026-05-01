import re
import datetime
import openpyxl
from docx import Document
from pathlib import Path

def format_value(placeholder_text, raw_value):
    """Formats currency for [$...] placeholders, dates for datetime values."""
    if isinstance(raw_value, (datetime.datetime, datetime.date)):
        return raw_value.strftime("%m/%d/%Y")

    is_dollar = bool(re.match(r"^\[\$", placeholder_text.strip()))
    if not is_dollar:
        return str(raw_value)
    try:
        num = float(str(raw_value).replace(",", "").replace("$", ""))
    except ValueError:
        return str(raw_value)
    return f"${num:,.2f}"

def read_excel_mappings(excel_path):
    """Reads placeholders and values from an Excel file."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = None
    # Look for a sheet with 'field' in the name
    for name in wb.sheetnames:
        if "field" in name.lower():
            ws = wb[name]
            break
    
    if ws is None:
        raise ValueError(f"No sheet with 'field' in name found. Sheets: {wb.sheetnames}")

    mappings = []
    skipped = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        if len(row) < 3:
            continue
        field_cell, value_cell = row[1], row[2]
        if not field_cell or not str(field_cell).strip():
            continue
        field_raw = str(field_cell).strip()
        if not re.search(r"\[", field_raw):
            continue
        if not value_cell or not str(value_cell).strip():
            skipped.append(field_raw)
            continue
        mappings.append((field_raw, format_value(field_raw, str(value_cell).strip())))

    # Sort by length descending to prevent partial replacements (e.g., [FIELD] vs [FIELD_1])
    mappings.sort(key=lambda x: len(x[0]), reverse=True)
    return mappings, skipped

def read_excel_records(excel_path, sheet_name=None, placeholder_row=1, header_row=2, data_start_row=3):
    """Reads a tabular export sheet where one row in the header block lists
    bracketed placeholders. Returns (records, placeholder_columns, header_row_values).

    records: list of dicts shaped {'mappings': [(placeholder, value), ...], 'row': raw_row_tuple}
    placeholder_columns: list of (col_index, placeholder_text)
    header_row_values: tuple of human-readable column headers from `header_row`
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    ws = None
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        # Default to a sheet whose name starts with "Export" but isn't a pivot
        for name in wb.sheetnames:
            n = name.lower().strip()
            if n.startswith("export") and "pivot" not in n:
                ws = wb[name]
                break
        if ws is None:
            ws = wb[wb.sheetnames[0]]

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < data_start_row:
        raise ValueError(f"Sheet '{ws.title}' has too few rows for bulk fill.")

    placeholder_row_vals = rows[placeholder_row - 1]
    header_row_vals = rows[header_row - 1]

    placeholder_columns = []
    for idx, cell in enumerate(placeholder_row_vals):
        if cell is None:
            continue
        text = str(cell).strip()
        if not text or "[" not in text or "]" not in text:
            continue
        placeholder_columns.append((idx, text))

    if not placeholder_columns:
        raise ValueError(
            f"No bracketed placeholders found in row {placeholder_row} of sheet '{ws.title}'."
        )

    records = []
    for raw_row in rows[data_start_row - 1:]:
        if not any(c is not None and str(c).strip() != "" for c in raw_row):
            continue

        mappings = []
        for col_idx, placeholder in placeholder_columns:
            if col_idx >= len(raw_row):
                continue
            val = raw_row[col_idx]
            if val is None:
                continue
            s = str(val).strip()
            if not s:
                continue
            mappings.append((placeholder, format_value(placeholder, val)))

        if not mappings:
            continue

        mappings.sort(key=lambda x: len(x[0]), reverse=True)
        records.append({"mappings": mappings, "row": raw_row})

    return records, placeholder_columns, header_row_vals

def read_excel_records_column_oriented(excel_path, sheet_name=None):
    """Reads a column-oriented sheet where column B holds field names
    (placeholders like [Procedure] or labels like 'Patient Name') and each
    subsequent column (C onward) is one record.

    Returns list of dicts: {mappings, patient_name, procedure, dispute_id}.
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    ws = None
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        for name in wb.sheetnames:
            n = name.lower().strip()
            if "field" in n and ("replace" in n or "fill" in n or "enter" in n):
                ws = wb[name]
                break
        if ws is None:
            for name in wb.sheetnames:
                if "field" in name.lower():
                    ws = wb[name]
                    break
        if ws is None:
            raise ValueError(f"No field sheet found. Sheets: {wb.sheetnames}")

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"Sheet '{ws.title}' is empty.")

    max_cols = max((len(r) for r in rows), default=0)
    if max_cols < 3:
        raise ValueError(
            f"Sheet '{ws.title}' has too few columns for column-oriented fill."
        )

    placeholder_rows = []
    patient_row = None
    procedure_row = None
    for idx, row in enumerate(rows):
        if len(row) < 2 or row[1] is None:
            continue
        label = str(row[1]).strip()
        if not label:
            continue
        if "[" in label and "]" in label:
            placeholder_rows.append((idx, label))
            if re.sub(r"\s+", "", label).lower() == "[procedure]":
                procedure_row = idx
        if label.lower() == "patient name":
            patient_row = idx

    if not placeholder_rows:
        raise ValueError(
            f"No bracketed placeholders found in column B of sheet '{ws.title}'."
        )

    records = []
    for col in range(2, max_cols):
        mappings = []
        for r_idx, placeholder in placeholder_rows:
            r = rows[r_idx]
            if col >= len(r):
                continue
            val = r[col]
            if val is None:
                continue
            s = str(val).strip()
            if not s:
                continue
            mappings.append((placeholder, format_value(placeholder, val)))

        if not mappings:
            continue

        mappings.sort(key=lambda x: len(x[0]), reverse=True)

        patient_name = None
        if patient_row is not None and col < len(rows[patient_row]):
            v = rows[patient_row][col]
            if v is not None and str(v).strip():
                patient_name = str(v).strip()

        procedure = None
        if procedure_row is not None and col < len(rows[procedure_row]):
            v = rows[procedure_row][col]
            if v is not None and str(v).strip():
                procedure = str(v).strip()

        dispute_id = None
        for ph, vv in mappings:
            if re.sub(r"\s+", "", ph).lower().strip("[]") == "disputeid":
                dispute_id = vv
                break

        records.append({
            "mappings": mappings,
            "patient_name": patient_name,
            "procedure": procedure,
            "dispute_id": dispute_id,
        })

    return records


def is_scs_procedure(procedure_text):
    """True if the procedure text references a Spinal Cord Stimulator."""
    if not procedure_text:
        return False
    s = str(procedure_text)
    if "spinal cord stimulator" in s.lower():
        return True
    if re.search(r"\bSCS\b", s):
        return True
    return False


def safe_filename_part(value, fallback="record"):
    """Sanitizes a value for use in a filename."""
    if value is None:
        return fallback
    if isinstance(value, (datetime.datetime, datetime.date)):
        s = value.strftime("%Y-%m-%d")
    else:
        s = str(value)
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s).strip("_")
    return s or fallback

def get_all_runs(paragraph):
    """Helper to extract all runs from a paragraph, including those in hyperlinks."""
    from docx.text.run import Run
    runs = []
    for child in paragraph._p:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "r":
            runs.append(Run(child, paragraph))
        elif tag == "hyperlink":
            for r_elem in child:
                r_tag = r_elem.tag.split("}")[-1] if "}" in r_elem.tag else r_elem.tag
                if r_tag == "r":
                    runs.append(Run(r_elem, paragraph))
    return runs

def replace_in_paragraph(paragraph, mappings):
    """Replaces placeholders across multiple runs in a paragraph."""
    runs = get_all_runs(paragraph)
    if not runs:
        return 0
    
    char_map = []
    for i, run in enumerate(runs):
        for ch in run.text:
            char_map.append((ch, i))
    
    if not char_map:
        return 0
        
    full_text = "".join(c for c, _ in char_map)
    full_text_low = full_text.lower()
    count = 0
    
    for placeholder, value in mappings:
        search_str = placeholder.lower()
        start_idx = 0
        while True:
            idx = full_text_low.find(search_str, start_idx)
            if idx == -1:
                break
            
            end_idx = idx + len(search_str)
            # Find which run this placeholder starts in
            run_index = char_map[idx][1]
            
            # Update char_map and full_text to reflect the replacement
            char_map = char_map[:idx] + [(ch, run_index) for ch in value] + char_map[end_idx:]
            full_text = full_text[:idx] + value + full_text[end_idx:]
            full_text_low = full_text_low[:idx] + value.lower() + full_text_low[end_idx:]
            
            count += 1
            start_idx = idx + len(value)
            
    if count == 0:
        return 0
        
    # Reconstruct run texts
    run_texts = {i: [] for i in range(len(runs))}
    for ch, i in char_map:
        run_texts[i].append(ch)
        
    for i, run in enumerate(runs):
        run.text = "".join(run_texts[i])
        
    return count

def fill_document(word_file, mappings):
    """Processes a Word document and replaces all placeholders."""
    doc = Document(word_file)
    total_replacements = 0
    
    # Process paragraphs
    for p in doc.paragraphs:
        total_replacements += replace_in_paragraph(p, mappings)
        
    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    total_replacements += replace_in_paragraph(p, mappings)
                    
    # Process headers and footers
    for section in doc.sections:
        headers_footers = [
            section.header, section.footer,
            section.even_page_header, section.even_page_footer,
            section.first_page_header, section.first_page_footer
        ]
        for hf in headers_footers:
            if hf:
                for p in hf.paragraphs:
                    total_replacements += replace_in_paragraph(p, mappings)
                for table in hf.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                total_replacements += replace_in_paragraph(p, mappings)

    return doc, total_replacements
