import os
import glob
import re
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# ==========================================
# CONFIGURATION & STYLING SETTINGS
# ==========================================
INPUT_FOLDER = 'input_csvs'
OUTPUT_FOLDER = 'output_pdfs'

# Sub-folders for organization
FULL_REPORTS_FOLDER = os.path.join(OUTPUT_FOLDER, 'Full_Reports')
INDIVIDUAL_RECORDS_ROOT = os.path.join(OUTPUT_FOLDER, 'Individual_Records')

# --- Table Visuals ---
FONT_NAME_BOLD = 'Helvetica-Bold'
FONT_NAME_REGULAR = 'Helvetica'
FONT_SIZE = 10
HEADER_BG_COLOR = colors.grey
HEADER_TEXT_COLOR = colors.whitesmoke
ROW_BG_COLOR_1 = colors.whitesmoke
ROW_BG_COLOR_2 = colors.beige
GRID_COLOR = colors.black
PADDING = 8  # Cell padding

def sanitize_filename(name):
    """Removes or replaces characters that are unsafe for filenames."""
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).replace(" ", "_")

def calculate_page_size(data):
    """
    Dynamically calculates the page dimensions based on content.
    This ensures the PDF is scalable for very wide or very long CSVs.
    """
    if not data:
        return (8.5 * inch, 11 * inch), [100]

    num_cols = len(data[0])
    num_rows = len(data)
    
    col_widths = []
    for i in range(num_cols):
        max_len = 0
        for row in data:
            if i < len(row):
                cell_content = str(row[i])
                max_len = max(max_len, len(cell_content))
        # Logic: Estimate width based on character count.
        # min width of 100 points, scaling up by ~7 points per character.
        col_widths.append(max(max_len * 7, 100)) 

    # Add margins to the total dimensions
    total_width = sum(col_widths) + 100 
    total_height = (num_rows * (FONT_SIZE + PADDING * 2)) + 150 

    # Return (Width, Height) and the list of individual column widths
    return (max(total_width, 8.5 * inch), max(total_height, 11 * inch)), col_widths

def generate_pdf(data, output_path, title_text):
    """
    Core engine for PDF generation using ReportLab Platypus.
    Modify TableStyle here to change global styling.
    """
    try:
        page_size, col_widths = calculate_page_size(data)
        
        # Setup document template
        doc = SimpleDocTemplate(
            output_path, 
            pagesize=page_size, 
            rightMargin=40, 
            leftMargin=40, 
            topMargin=40, 
            bottomMargin=40
        )
        
        elements = []
        styles = getSampleStyleSheet()
        
        # Add Title
        title = Paragraph(title_text, styles['Title'])
        elements.append(title)
        elements.append(Spacer(1, 0.2 * inch))

        # Create the Table object
        t = Table(data, colWidths=col_widths)
        
        # --- TABLE STYLING SECTION ---
        # Modify these values to change the look of your PDFs
        table_style = TableStyle([
            # Header Styling (Row 0)
            ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG_COLOR),
            ('TEXTCOLOR', (0, 0), (-1, 0), HEADER_TEXT_COLOR),
            ('FONTNAME', (0, 0), (-1, 0), FONT_NAME_BOLD),
            
            # General Alignment and Font
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), FONT_SIZE),
            
            # Grid and Borders
            ('GRID', (0, 0), (-1, -1), 0.5, GRID_COLOR),
            
            # Padding
            ('TOPPADDING', (0, 0), (-1, -1), PADDING),
            ('BOTTOMPADDING', (0, 0), (-1, -1), PADDING),
            ('LEFTPADDING', (0, 0), (-1, -1), PADDING),
            ('RIGHTPADDING', (0, 0), (-1, -1), PADDING),
        ])
        
        # Alternating Background Colors for Data Rows
        for i in range(1, len(data)):
            bg_color = ROW_BG_COLOR_1 if i % 2 == 0 else ROW_BG_COLOR_2
            table_style.add('BACKGROUND', (0, i), (-1, i), bg_color)
            
        # Bold the first column (The 'Schema' labels)
        table_style.add('FONTNAME', (0, 0), (0, -1), FONT_NAME_BOLD)
            
        t.setStyle(table_style)
        elements.append(t)
        
        # Render the PDF
        doc.build(elements)
    except Exception as e:
        print(f"Failed to generate PDF at {output_path}: {e}")

def convert_file_to_pdf(file_path):
    print(f"\nAnalyzing: {file_path}")
    try:
        # Determine how to read the file based on extension
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        else:
            print(f"Unsupported file type: {ext}")
            return
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    if df.empty:
        print("File is empty. Skipping.")
        return

    filename = os.path.basename(file_path)
    name_no_ext = os.path.splitext(filename)[0]

    # --- 1. FULL REPORT ---
    df_transposed = df.transpose()
    df_transposed.reset_index(inplace=True)
    
    main_headers = ["Schema / Field"] + [f"Record {i+1}" for i in range(len(df))]
    main_data = [main_headers] + df_transposed.values.astype(str).tolist()
    
    # Save in the dedicated Full_Reports folder
    main_output_path = os.path.join(FULL_REPORTS_FOLDER, f"{name_no_ext}_FULL_REPORT.pdf")
    generate_pdf(main_data, main_output_path, f"Complete Data: {filename}")
    print(f"  [✔] Full Report Saved: {main_output_path}")

    # --- 2. INDIVIDUAL RECORDS ---
    # Save in a subfolder inside Individual_Records
    records_subfolder = os.path.join(INDIVIDUAL_RECORDS_ROOT, f"{name_no_ext}_individual_records")
    if not os.path.exists(records_subfolder):
        os.makedirs(records_subfolder)
    
    schemas = df.columns.tolist()
    
    naming_cols = [col for col in df.columns if any(x in str(col).lower() for x in ['name', 'candidate', 'id'])]
    
    print(f"  [→] Generating {len(df)} individual PDFs...")
    
    for idx, row in df.iterrows():
        name_parts = [str(row[col]) for col in naming_cols[:2]] 
        if not name_parts or all(not p or p.lower() == 'nan' for p in name_parts):
            name_parts = [f"Record_{idx+1}"]
        
        sanitized_name = sanitize_filename("_".join(name_parts))
        final_pdf_name = f"{sanitized_name}_{name_no_ext}.pdf"
        
        record_table_data = [["Field Name", "Value"]]
        for schema in schemas:
            record_table_data.append([str(schema), str(row[schema])])
            
        record_output_path = os.path.join(records_subfolder, final_pdf_name)
        generate_pdf(record_table_data, record_output_path, f"Record Details: {sanitized_name.replace('_', ' ')}")
        
    print(f"  [✔] Individual records saved in: {records_subfolder}")

def main():
    # Setup directories
    folders_to_create = [
        INPUT_FOLDER, 
        OUTPUT_FOLDER, 
        FULL_REPORTS_FOLDER, 
        INDIVIDUAL_RECORDS_ROOT
    ]
    for folder in folders_to_create:
        if not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)

    # Support both CSV and Excel files
    extensions = ['*.csv', '*.xlsx', '*.xls']
    files = []
    for ext in extensions:
        files.extend(glob.glob(os.path.join(INPUT_FOLDER, ext)))
    
    if not files:
        print(f"Notice: No supported files found in '{INPUT_FOLDER}'.")
        return

    for f in files:
        convert_file_to_pdf(f)
        
    print("\nProcessing Complete.")

if __name__ == "__main__":
    main()