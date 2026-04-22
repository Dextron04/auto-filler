import re
import openpyxl
from docx import Document
from pathlib import Path

def format_value(placeholder_text, raw_value):
    """Formats values as currency if the placeholder starts with [$]."""
    is_dollar = bool(re.match(r"^\[\$", placeholder_text.strip()))
    if not is_dollar:
        return raw_value
    try:
        # Handle cases where raw_value might already have symbols or commas
        num = float(str(raw_value).replace(",", "").replace("$", ""))
    except ValueError:
        return raw_value
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
