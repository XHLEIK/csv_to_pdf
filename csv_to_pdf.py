import os
import glob
import re
import pandas as pd
import urllib.request
import shutil
import tempfile
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT, TA_JUSTIFY

# ==========================================
# CONFIGURATION & STYLING SETTINGS
# ==========================================
INPUT_FOLDER = 'input_csvs'
OUTPUT_FOLDER = 'output_pdfs'

# Sub-folders
FULL_REPORTS_FOLDER = os.path.join(OUTPUT_FOLDER, 'Full_Reports')
INDIVIDUAL_RECORDS_ROOT = os.path.join(OUTPUT_FOLDER, 'Individual_Records')

# --- Table Visuals ---
FONT_NAME_BOLD = 'Helvetica-Bold'
FONT_NAME_REGULAR = 'Helvetica'
FONT_SIZE = 10
HEADER_BG_COLOR = colors.white
TABLE_HEADER_BG = colors.Color(0.9, 0.9, 0.9) # Light Grey
GRID_COLOR = colors.black
PADDING = 6

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).replace(" ", "_")

def download_signature_from_gdrive(url):
    """Downloads an image from a Google Drive URL to a temp file, trying multiple strategies."""
    try:
        url = str(url).strip()
        if not url or url.lower() == 'nan' or not url.startswith('http'):
            return None
            
        file_id = None
        match = re.search(r'(?:id=|[d]/|file/d/|open\?id=)([a-zA-Z0-9_-]+)', url)
        if match:
            file_id = match.group(1)
        
        if not file_id:
            return None
            
        # Strategy 1: uc?export=download
        download_urls = [
            f'https://drive.google.com/uc?export=download&id={file_id}',
            f'https://drive.google.com/uc?id={file_id}&export=download',
            f'https://drive.google.com/thumbnail?id={file_id}&sz=w1000' # Thumbnail can sometimes bypass restrictions
        ]
        
        tmp_path = None
        for d_url in download_urls:
            try:
                # Create temp file
                tmp_fd, current_tmp = tempfile.mkstemp(suffix='.jpg')
                os.close(tmp_fd)
                
                req = urllib.request.Request(d_url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=10) as response:
                    content = response.read()
                    
                    # Handle confirmation page
                    if b'confirm=' in content:
                        confirm_token = re.search(r'confirm=([a-zA-Z0-9_-]+)', content.decode('utf-8', errors='ignore'))
                        if confirm_token:
                            d_url_confirm = d_url + f"&confirm={confirm_token.group(1)}"
                            req = urllib.request.Request(d_url_confirm, headers={'User-Agent': 'Mozilla/5.0'})
                            with urllib.request.urlopen(req, timeout=10) as response2:
                                content = response2.read()
                    
                    # Check if valid image
                    if (content.startswith(b'\xff\xd8') or content.startswith(b'\x89PNG') or 
                        content.startswith(b'GIF87a') or content.startswith(b'GIF89a')):
                        with open(current_tmp, 'wb') as out_file:
                            out_file.write(content)
                        tmp_path = current_tmp
                        break # Success
                    else:
                        os.remove(current_tmp)
            except Exception:
                continue
                
        return tmp_path
    except Exception as e:
        print(f"Error in download_signature_from_gdrive: {e}")
        return None

def get_clean_title(filename):
    """Extracts 'PRE-EXAM REPORT' from 'PRE-EXAM REPORT (Responses)...'"""
    name = os.path.splitext(filename)[0]
    if '(' in name:
        name = name.split('(')[0]
    return name.strip().upper()

def calculate_col_widths(data, page_width, min_width=50):
    """
    Calculates column widths based on content, fitting within page_width.
    Used mainly for the Full Report (Matrix view).
    """
    available_width = page_width - 80 # Margins
    num_cols = len(data[0])
    
    col_widths = []
    for i in range(num_cols):
        max_char = 0
        for row in data:
            if i < len(row):
                # Data here might be Paragraphs or strings depending on usage
                # For full report we currently use strings.
                cell_content = str(row[i])
                max_char = max(max_char, len(cell_content))
        col_widths.append(max(max_char * 6, min_width))
        
    total_est = sum(col_widths)
    if total_est < available_width:
        extra = (available_width - total_est) / num_cols
        col_widths = [w + extra for w in col_widths]
        
    return col_widths

def create_header_flowables(report_title):
    styles = getSampleStyleSheet()
    
    style_main_header = ParagraphStyle(
        'MainHeader',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=14,
        alignment=TA_CENTER,
        spaceAfter=6,
        textColor=colors.black
    )
    
    style_sub_header = ParagraphStyle(
        'SubHeader',
        parent=styles['Heading2'],
        fontName='Helvetica',
        fontSize=12,
        alignment=TA_CENTER,
        spaceAfter=6,
        textColor=colors.black
    )
    
    style_report_title = ParagraphStyle(
        'ReportTitle',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=11,
        alignment=TA_CENTER,
        spaceAfter=20,
        textColor=colors.black
    )

    elements = []
    elements.append(Paragraph("ARUNACHAL PRADESH PUBLIC SERVICE COMMISSION", style_main_header))
    elements.append(Paragraph("<u>OBSERVER REPORT</u>", style_sub_header))
    elements.append(Paragraph(report_title, style_report_title))
    return elements

def create_footer_flowables(observer_info, page_width):
    """
    Creates the post-table section with refined alignment for A4 Portrait.
    """
    styles = getSampleStyleSheet()
    
    style_date = ParagraphStyle(
        'FooterDate', 
        parent=styles['Normal'], 
        alignment=TA_LEFT, 
        leading=18, 
        fontSize=10,
        fontName='Helvetica-Bold'
    )

    style_label = ParagraphStyle(
        'LabelStyle',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica-Bold',
        alignment=TA_LEFT
    )

    style_value = ParagraphStyle(
        'ValueStyle',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica',
        alignment=TA_LEFT
    )

    style_value_large = ParagraphStyle(
        'ValueStyleLarge',
        parent=styles['Normal'],
        fontSize=12,
        fontName='Helvetica-Bold',
        alignment=TA_LEFT
    )
    
    # 1. Signature Handling
    sig_path = observer_info.get('signature_path')
    sig_img = None
    if sig_path and os.path.exists(sig_path):
        try:
            from PIL import Image
            img = Image.open(sig_path)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            fd, new_sig_path = tempfile.mkstemp(suffix='.png')
            os.close(fd)
            img.save(new_sig_path, format='PNG')
            sig_img = f'<img src="{new_sig_path}" width="80" height="30" valign="bottom" />'
        except:
            sig_img = None

    # Helper to create underlined cell content
    def get_underlined_para(text, style, min_width_chars=40):
        # We use a fixed-width underline by padding with non-breaking spaces
        padding = "&nbsp;" * max(0, min_width_chars - len(str(text)))
        return Paragraph(f"<u>&nbsp;{text}{padding}</u>", style)

    name_val = observer_info.get('name', '')
    venue_val = observer_info.get('venue', '')
    mobile_val = observer_info.get('mobile', '')
    
    # Constructing the Inner Table for Right Side
    # Column 1: Labels, Column 2: Underlined Values
    inner_data = [
        [Paragraph("Signature of Observer:", style_label), get_underlined_para(sig_img if sig_img else "", style_value, 35)],
        [Paragraph("Name:", style_label), get_underlined_para(name_val, style_value_large, 35)],
        [Paragraph("Mobile Number:", style_label), get_underlined_para(mobile_val, style_value, 35)],
        [Paragraph("Exam Venue:", style_label), get_underlined_para(venue_val, style_value_large, 35)]
    ]
    
    # Widths for inner table (approx 300 points total)
    t_inner = Table(inner_data, colWidths=[110, 190])
    t_inner.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))

    date_text = f"<br/><br/><br/><br/><b>Date:</b> {observer_info.get('date', '')}"
    p_date = Paragraph(date_text, style_date)
    
    # Main Footer Layout
    usable_width = page_width - 80
    col1_w = usable_width * 0.30
    col2_w = usable_width * 0.70
    
    footer_data = [[p_date, t_inner]]
    t_footer = Table(footer_data, colWidths=[col1_w, col2_w])
    t_footer.setStyle(TableStyle([
        ('VALIGN', (0,0), (0,0), 'BOTTOM'),
        ('VALIGN', (1,0), (1,0), 'TOP'),
        ('LEFTPADDING', (0,0), (0,0), 0),
        ('ALIGN', (1,0), (1,0), 'RIGHT'), # Align the inner table container to the right
    ]))
    
    return [Spacer(1, 0.4*inch), t_footer]

def get_wrapped_text(text, style=None):
    """Wraps text in a Paragraph object to ensure it wraps inside table cells."""
    if style is None:
        style = getSampleStyleSheet()['Normal']
        style.fontSize = FONT_SIZE
        style.leading = FONT_SIZE + 4
    
    # Clean text to prevent XML errors in ReportLab
    safe_text = str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br/>')
    return Paragraph(safe_text, style)

def generate_pdf_report(data, output_path, report_title, observer_info=None, is_individual=False):
    try:
        # Determine Page Size and Column Widths
        if is_individual:
            # STRICT A4 PORTRAIT
            page_width, page_height = A4
            
            # Margins: 40 points left, 40 points right
            available_width = page_width - 80 
            
            # Define exact column widths for A4
            # Sl No: Small (~8%)
            # Particulars: Wide (~57%)
            # Response: Medium (~35%)
            
            w_sl = available_width * 0.08
            w_part = available_width * 0.57
            w_resp = available_width * 0.35
            
            col_widths = [w_sl, w_part, w_resp]
            
            # Ensure data cells are wrapped Paragraphs, NOT strings
            # This logic happens inside convert_file_to_pdf now, or we can transform here if needed.
            # But simpler to pass 'flowables' in 'data' from the caller.
            # We will assume 'data' contains Paragraph objects for text fields.
            
        else:
            # Full Report (Matrix) - Dynamic Size
            page_width = 8.5 * inch
            page_height = 11 * inch
            
            # Recalculate widths based on raw data (assuming strings for full report for now or basic paragraphs)
            # For Full Report, we usually want strings or simple wrap.
            col_widths = calculate_col_widths(data, page_width)
            total_w = sum(col_widths) + 80
            if total_w > page_width:
                page_width = total_w
            
            # Estimate height
            est_height = len(data) * 30 + 400
            if est_height > page_height:
                page_height = est_height

        doc = SimpleDocTemplate(
            output_path, 
            pagesize=(page_width, page_height),
            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40
        )
        
        elements = []
        elements.extend(create_header_flowables(report_title))
        
        # Create Table
        t = Table(data, colWidths=col_widths, repeatRows=1)
        
        # Style
        style_list = [
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BG),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('FONTNAME', (0, 0), (-1, 0), FONT_NAME_BOLD),
            ('FONTSIZE', (0, 0), (-1, -1), FONT_SIZE),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'), # Important for wrapped text
            ('GRID', (0, 0), (-1, -1), 0.5, GRID_COLOR),
            ('TOPPADDING', (0, 0), (-1, -1), PADDING),
            ('BOTTOMPADDING', (0, 0), (-1, -1), PADDING),
            ('LEFTPADDING', (0, 0), (-1, -1), PADDING),
            ('RIGHTPADDING', (0, 0), (-1, -1), PADDING),
        ]
        
        # Center "Sl. No." column alignment
        style_list.append(('ALIGN', (0,0), (0,-1), 'CENTER'))
        
        t.setStyle(TableStyle(style_list))
        elements.append(t)
        
        if observer_info:
            elements.extend(create_footer_flowables(observer_info, page_width))
            
        doc.build(elements)
        
    except Exception as e:
        print(f"Error generating PDF {output_path}: {e}")

def clean_particulars(text):
    """Removes leading numbers like '1.', '2. ', '10) ' from questions."""
    return re.sub(r'^\d+[\s.)-]*', '', str(text)).strip()

def convert_file_to_pdf(file_path):
    print(f"\nAnalyzing: {file_path}")
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        else:
            return
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    if df.empty:
        return

    filename = os.path.basename(file_path)
    name_no_ext = os.path.splitext(filename)[0]
    clean_title = get_clean_title(filename)

    # Prepare Styles for Table Cells
    styles = getSampleStyleSheet()
    cell_style = styles['Normal']
    cell_style.fontSize = 10
    cell_style.leading = 12

    # ==========================================
    # 1. FULL REPORT
    # ==========================================
    df_transposed = df.transpose()
    df_transposed.reset_index(inplace=True)
    
    full_data_rows = df_transposed.values.astype(str).tolist()
    
    matrix_table = []
    # Header (Strings are fine for header)
    matrix_table.append(["Sl. No.", "Particulars"] + [f"Rec {i+1}" for i in range(len(df))])
    
    for i, row in enumerate(full_data_rows):
        p_clean = clean_particulars(row[0])
        # We can use Paragraphs for Full Report too if we want wrapping there
        # But user emphasized Individual Report A4.
        # Let's keep Full Report simple strings for now to maintain grid look, 
        # or use Paragraphs if content is huge. Let's use strings for matrix to save processing.
        matrix_table.append([str(i+1), p_clean] + row[1:])

    full_output_path = os.path.join(FULL_REPORTS_FOLDER, f"{name_no_ext}_FULL_REPORT.pdf")
    generate_pdf_report(matrix_table, full_output_path, clean_title, observer_info=None)
    print(f"  [✔] Full Report Saved: {full_output_path}")

    # ==========================================
    # 2. INDIVIDUAL RECORDS
    # ==========================================
    records_subfolder = os.path.join(INDIVIDUAL_RECORDS_ROOT, f"{name_no_ext}_individual_records")
    if not os.path.exists(records_subfolder):
        os.makedirs(records_subfolder)
        
    schemas = df.columns.tolist()
    cols_map = {c.lower().strip(): c for c in schemas}
    
    def find_key(keywords):
        for k in keywords:
            for col_lower in cols_map:
                if k in col_lower:
                    return cols_map[col_lower]
        return None

    col_name = find_key(['name :', 'candidate name', 'name of candidate', 'name'])
    col_venue = find_key(['name of exam venue', 'venue', 'center'])
    col_date = find_key(['timestamp', 'date'])
    col_signature = find_key(['signature', 'upload signature', 'observer signature'])
    col_mobile = find_key(['mobile', 'phone', 'contact'])
    
    naming_cols = [col for col in df.columns if any(x in str(col).lower() for x in ['name', 'candidate', 'id'])]

    print(f"  [→] Generating {len(df)} individual PDFs...")
    
    for idx, row in df.iterrows():
        footer_name = str(row[col_name]) if col_name else ""
        footer_venue = str(row[col_venue]) if col_venue else ""
        footer_date = str(row[col_date]) if col_date else ""
        footer_mobile = str(row[col_mobile]) if col_mobile else ""
        
        # Clean mobile if it is NaN
        if footer_mobile.lower() == 'nan': footer_mobile = ""

        # Handle Signature
        sig_url = str(row[col_signature]) if col_signature else None
        sig_path = download_signature_from_gdrive(sig_url)
        
        observer_info = {
            'name': footer_name, 
            'venue': footer_venue, 
            'date': footer_date, 
            'signature_path': sig_path,
            'mobile': footer_mobile
        }

        # Header Row (Strings)
        table_data = [["Sl. No.", "Particulars", "Response"]]
        
        counter = 1
        for schema in schemas:
            if schema in [col_name, col_venue, col_date, col_signature, col_mobile]:
                continue
                
            p_clean = clean_particulars(schema)
            val = str(row[schema])
            if val.lower() == 'nan': val = ""
            
            # Specific requirement: Certification question always shows "Yes"
            if "I hereby certify" in p_clean:
                val = "Yes"
            
            # WRAP CONTENT IN PARAGRAPHS FOR INDIVIDUAL REPORTS
            # This enables word wrapping in A4 cells
            p_sl = str(counter)
            p_particular = get_wrapped_text(p_clean, cell_style)
            p_response = get_wrapped_text(val, cell_style)
            
            table_data.append([p_sl, p_particular, p_response])
            counter += 1

        name_parts = [str(row[col]) for col in naming_cols[:2]] 
        if not name_parts or all(not p or p.lower() == 'nan' for p in name_parts):
            name_parts = [f"Record_{idx+1}"]
        
        sanitized_name = sanitize_filename("_".join(name_parts))
        final_pdf_name = f"{sanitized_name}_{name_no_ext}.pdf"
        
        out_path = os.path.join(records_subfolder, final_pdf_name)
        generate_pdf_report(table_data, out_path, clean_title, observer_info, is_individual=True)

    print(f"  [✔] Individual records saved in: {records_subfolder}")

def main():
    for folder in [INPUT_FOLDER, OUTPUT_FOLDER, FULL_REPORTS_FOLDER, INDIVIDUAL_RECORDS_ROOT]:
        if not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)

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
