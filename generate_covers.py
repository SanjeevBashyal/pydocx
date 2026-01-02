import pandas as pd
from docx import Document
from docx2pdf import convert
from pypdf import PdfWriter
import os
import sys

# Configuration
COVER_DATA_FILE = 'cover_data.xlsx'
TEMPLATE_MAIN = 'Annex cover main.docx'
TEMPLATE_SUB = 'Annex cover sub main.docx'

def replace_text_in_paragraph(paragraph, key, value):
    """
    Replace occurrences of 'key' with 'value' in a paragraph.
    First tries to find the key in individual runs to preserve formatting.
    Falls back to text replacement if split across runs (may lose formatting).
    """
    if key in paragraph.text:
        replaced_in_run = False
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(key, value)
                replaced_in_run = True
        
        # If the key was found in paragraph.text but not in any single run,
        # it means the key is split across runs (e.g. bold/unbold boundaries).
        # In this case, we have to replace the whole text, which loses formatting.
        if not replaced_in_run:
            # Alternative: smarter run merging, but for now simple fallback.
            paragraph.text = paragraph.text.replace(key, value)

def process_template(row, template_path, output_docx_path):
    doc = Document(template_path)
    
    # Iterate over columns in the row
    for key, value in row.items():
        search_text = str(key)
        # Handle NaN/None
        replace_text = str(value) if pd.notna(value) else ""
        
        # Check paragraphs in body
        for p in doc.paragraphs:
            replace_text_in_paragraph(p, search_text, replace_text)
            
        # Check tables
        for table in doc.tables:
            for row_obj in table.rows:
                for cell in row_obj.cells:
                    for p in cell.paragraphs:
                        replace_text_in_paragraph(p, search_text, replace_text)
    
    doc.save(output_docx_path)

def main():
    if not os.path.exists(COVER_DATA_FILE):
        print(f"Error: {COVER_DATA_FILE} not found.")
        return
    # Templates checked later or assume existence for now

    # 1. Read Data
    try:
        # Read Excel
        df = pd.read_excel(COVER_DATA_FILE)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    print("Headers detected:", df.columns.tolist())
    
    # Check if <FILE-NAME> column exists
    file_name_col = '<FILE-NAME>'
    if file_name_col not in df.columns:
        # Try to find a column that looks like it (case insensitive or stripped)
        match = next((c for c in df.columns if c.strip() == '<FILE-NAME>'), None)
        if match:
            file_name_col = match
        else:
            print(f"Error: Column '{file_name_col}' not found in CSV.")
            return

    # Create PDFs directory
    output_dir = 'PDFs'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")

    # 2. Generate DOCX and Convert to PDF
    print("Processing rows...")
    for index, row in df.iterrows():
        # Get filename
        file_name = row.get(file_name_col)
        if pd.isna(file_name) or str(file_name).strip() == "":
            print(f"  Warning: Row {index + 1} has empty {file_name_col}. Skipping.")
            continue
        
        file_name = str(file_name).strip()
        if not file_name.lower().endswith('.pdf'):
            file_name += '.pdf'
            
        target_pdf_path = os.path.join(output_dir, file_name)
        target_pdf_path = os.path.abspath(target_pdf_path)
        
        temp_docx = f"temp_cover_{index}.docx"
        
        # Determine template
        sub_comp_val = row.get('<SUB-COMPONENT>')
        if pd.isna(sub_comp_val) or str(sub_comp_val).strip() == "":
            selected_template = TEMPLATE_MAIN
        else:
            selected_template = TEMPLATE_SUB
            
        if not os.path.exists(selected_template):
            print(f"  Error: Template '{selected_template}' not found for row {index + 1}. Skipping.")
            continue
        
        try:
            print(f"  Processing row {index + 1}: {file_name}")
            print(f"    Using template: {selected_template}")
            process_template(row, selected_template, temp_docx)
            
            # Convert to PDF
            print(f"    Converting to {target_pdf_path}...")
            convert(os.path.abspath(temp_docx), target_pdf_path)
            
        except Exception as e:
            print(f"  Error processing row {index + 1}: {e}")
        finally:
            # Cleanup temp docx
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass

    print("Done.")

if __name__ == "__main__":
    main()
