from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from reportlab.lib.pagesizes import landscape, A0
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase.pdfmetrics import stringWidth
import os
import tempfile
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB  max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'xlsm', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def excel_to_pdf(excel_path, pdf_path):
    """Convert Excel to PDF with all columns on each page and repeating headers"""
    
    # Read the Excel or CSV file
    if excel_path.lower().endswith('.csv'):
        df = pd.read_csv(excel_path)
    else:
        df = pd.read_excel(excel_path)
    
    # Convert all data to strings and handle NaN values
    df = df.fillna('')
    df = df.astype(str)
    
    # Prepare data
    headers = df.columns.tolist()
    data_rows = df.values.tolist()
    num_columns = len(headers)
    num_rows = len(data_rows)
    
    if num_columns == 0 or num_rows == 0:
        raise ValueError("Excel file is empty")
    
    # Calculate optimal width for each column
    def calculate_optimal_width(header, column_data, min_width=40, max_width=200):
        header_font = 'Helvetica-Bold'
        data_font = 'Helvetica'
        header_size = 8
        data_size = 7
        
        header_width = stringWidth(str(header), header_font, header_size)
        max_content_width = header_width
        
        sample_size = min(100, len(column_data))
        for value in column_data[:sample_size]:
            content_width = stringWidth(str(value), data_font, data_size)
            max_content_width = max(max_content_width, content_width)
        
        optimal_width = max_content_width * 1.25
        optimal_width = max(min_width, min(optimal_width, max_width))
        
        return optimal_width
    
    # Calculate optimal widths
    col_widths = []
    for i, header in enumerate(headers):
        column_data = [row[i] for row in data_rows]
        width = calculate_optimal_width(header, column_data)
        col_widths.append(width)
    
    total_width_needed = sum(col_widths)
    
    # Determine page size
    a0_landscape_width = landscape(A0)[0]
    a0_landscape_height = landscape(A0)[1]
    margin_space = 72
    
    if total_width_needed + margin_space <= a0_landscape_width:
        page_size = landscape(A0)
    else:
        custom_width = total_width_needed + margin_space
        custom_height = a0_landscape_height
        page_size = (custom_width, custom_height)
    
    # Create paragraph styles
    centered_style = ParagraphStyle(
        'centered',
        alignment=TA_CENTER,
        fontSize=7,
        leading=9,
        wordWrap='CJK'
    )
    
    header_style = ParagraphStyle(
        'header',
        alignment=TA_CENTER,
        fontSize=8,
        leading=10,
        fontName='Helvetica-Bold',
        textColor=colors.whitesmoke,
        wordWrap='CJK'
    )
    
    # Convert data to Paragraphs
    wrapped_headers = [Paragraph(str(h), header_style) for h in headers]
    wrapped_data = []
    for row in data_rows:
        wrapped_row = [Paragraph(str(cell), centered_style) for cell in row]
        wrapped_data.append(wrapped_row)
    
    # Create PDF
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=page_size,
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )
    
    # Calculate rows per page
    available_height = page_size[1] - 72
    rows_per_page = max(15, int(available_height / 25))
    
    story = []
    
    # Process data in chunks
    for chunk_start in range(0, len(wrapped_data), rows_per_page):
        chunk_end = min(chunk_start + rows_per_page, len(wrapped_data))
        chunk_data = [wrapped_headers] + wrapped_data[chunk_start:chunk_end]
        
        # Create table with optimal column widths
        table = Table(chunk_data, colWidths=col_widths, repeatRows=1)
        
        # Apply styling
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ])
        
        table.setStyle(table_style)
        story.append(table)
        
        if chunk_end < len(wrapped_data):
            story.append(PageBreak())
    
    # Build the PDF
    doc.build(story)
    
    # Return metadata
    num_pages = (len(wrapped_data) + rows_per_page - 1) // rows_per_page
    return {
        'rows': num_rows,
        'columns': num_columns,
        'pages': num_pages,
        'page_width': page_size[0],
        'page_height': page_size[1]
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Please upload .xlsx, .xls, or .xlsm'}), 400
    
    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{filename}")
        file.save(excel_path)
        
        # Generate PDF path
        pdf_filename = filename.rsplit('.', 1)[0] + '.pdf'
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{pdf_filename}")
        
        # Convert
        metadata = excel_to_pdf(excel_path, pdf_path)
        
        # Clean up Excel file
        os.remove(excel_path)
        
        # Store PDF path in session (simplified - in production use proper session management)
        return jsonify({
            'success': True,
            'filename': pdf_filename,
            'download_path': f"/download/{timestamp}_{pdf_filename}",
            'metadata': metadata
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(filepath):
        response = send_file(filepath, as_attachment=True, download_name=filename.split('_', 1)[1])
        # Schedule file deletion after download (simplified)
        return response
    else:
        return "File not found", 404

import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
