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
import subprocess
import shutil

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB  max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'xlsm', 'csv'}
EXCEL_ONLY_EXTENSIONS = {'xlsx', 'xls', 'xlsm'}  # For LibreOffice conversion (no CSV)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def allowed_excel_file(filename):
    """Check if file is an Excel file (not CSV) for LibreOffice conversion"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in EXCEL_ONLY_EXTENSIONS

def get_libreoffice_path():
    """Find LibreOffice executable path"""
    # Common paths for LibreOffice
    possible_paths = [
        # Linux
        '/usr/bin/libreoffice',
        '/usr/bin/soffice',
        '/usr/lib/libreoffice/program/soffice',
        # macOS
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        # Windows
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ]

    # Check if libreoffice is in PATH
    libreoffice_in_path = shutil.which('libreoffice') or shutil.which('soffice')
    if libreoffice_in_path:
        return libreoffice_in_path

    # Check common paths
    for path in possible_paths:
        if os.path.exists(path):
            return path

    return None

def get_libreoffice_python():
    """Find LibreOffice's Python interpreter (has UNO built-in)"""
    possible_paths = [
        # Linux
        '/usr/bin/python3',  # System python with uno package
        '/usr/lib/libreoffice/program/python',
        # macOS
        '/Applications/LibreOffice.app/Contents/Resources/python',
        # Windows
        r'C:\Program Files\LibreOffice\program\python.exe',
        r'C:\Program Files (x86)\LibreOffice\program\python.exe',
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    # Fall back to system python
    return shutil.which('python3') or shutil.which('python')

def is_libreoffice_available():
    """Check if LibreOffice is installed and available"""
    return get_libreoffice_path() is not None

def excel_to_pdf_libreoffice(excel_path, output_dir):
    """Convert Excel to PDF using LibreOffice with each sheet scaled to fit one page"""
    libreoffice_path = get_libreoffice_path()
    if not libreoffice_path:
        raise RuntimeError("LibreOffice is not installed or not found")

    # Determine output path
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

    # Create a Python script that uses UNO to scale sheets and export
    # This script will be run by LibreOffice's Python interpreter
    script_content = f'''#!/usr/bin/env python3
import sys
import os
import time
import socket
import subprocess

def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

def main():
    input_path = {repr(os.path.abspath(excel_path))}
    output_path = {repr(os.path.abspath(pdf_path))}
    port = find_free_port()

    # Start LibreOffice in listening mode
    lo_process = subprocess.Popen([
        {repr(libreoffice_path)},
        '--headless',
        '--invisible',
        '--nodefault',
        '--nolockcheck',
        '--nologo',
        '--nofirststartwizard',
        f'--accept=socket,host=localhost,port={{port}};urp;StarOffice.ServiceManager'
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    try:
        # Wait for LibreOffice to start
        time.sleep(2)

        import uno
        from com.sun.star.beans import PropertyValue

        # Connect to LibreOffice
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext)

        # Try to connect with retries
        ctx = None
        for attempt in range(10):
            try:
                ctx = resolver.resolve(
                    f"uno:socket,host=localhost,port={{port}};urp;StarOffice.ComponentContext")
                break
            except:
                time.sleep(1)

        if ctx is None:
            raise RuntimeError("Could not connect to LibreOffice")

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        # Open the document
        url = uno.systemPathToFileUrl(input_path)
        props = (PropertyValue(Name="Hidden", Value=True),)
        doc = desktop.loadComponentFromURL(url, "_blank", 0, props)

        if doc is None:
            raise RuntimeError("Could not open document")

        try:
            # Process each sheet
            sheets = doc.getSheets()
            for i in range(sheets.getCount()):
                sheet = sheets.getByIndex(i)

                # 1. Get used range to check if sheet has content
                cursor = sheet.createCursor()
                cursor.gotoStartOfUsedArea(False)
                cursor.gotoEndOfUsedArea(True)
                used_range = cursor.getRangeAddress()

                # 2. Skip empty sheets (prevents blank pages)
                if used_range.EndColumn < used_range.StartColumn:
                    continue

                # 3. Set print area to only the used range
                sheet.setPrintAreas([used_range])

                # 4. Configure page style for fit-width scaling
                style_name = sheet.PageStyle
                styles = doc.getStyleFamilies().getByName("PageStyles")
                page_style = styles.getByName(style_name)

                # 5. Set landscape orientation with A2 paper size
                page_style.IsLandscape = True
                page_style.Width = 59400   # A2 width: 594mm (1/100 mm units)
                page_style.Height = 42000  # A2 height: 420mm

                # 6. Fit entire sheet to 1 page
                page_style.ScaleToPagesX = 1
                page_style.ScaleToPagesY = 1

                # 7. Reduce margins for more content space (1/100 mm units)
                page_style.LeftMargin = 1000   # 10mm
                page_style.RightMargin = 1000  # 10mm
                page_style.TopMargin = 1000    # 10mm
                page_style.BottomMargin = 1000 # 10mm

            # Export to PDF
            output_url = uno.systemPathToFileUrl(output_path)
            export_props = (
                PropertyValue(Name="FilterName", Value="calc_pdf_Export"),
            )
            doc.storeToURL(output_url, export_props)
        finally:
            doc.close(True)

        # Shutdown LibreOffice
        desktop.terminate()

    finally:
        lo_process.terminate()
        lo_process.wait(timeout=10)

    if not os.path.exists(output_path):
        sys.exit(1)
    sys.exit(0)

if __name__ == "__main__":
    main()
'''

    # Write the script to a temp file
    script_path = os.path.join(tempfile.gettempdir(), f'lo_convert_{os.getpid()}.py')
    try:
        with open(script_path, 'w') as f:
            f.write(script_content)

        # Run the script with system Python (uno module available on Linux)
        python_path = get_libreoffice_python()
        env = os.environ.copy()
        env['PYTHONPATH'] = '/usr/lib/libreoffice/program'
        result = subprocess.run(
            [python_path, script_path],
            capture_output=True,
            text=True,
            timeout=300,
            env=env
        )

        if result.returncode != 0:
            # Fall back to simple conversion without scaling
            return excel_to_pdf_libreoffice_simple(excel_path, output_dir)

        if not os.path.exists(pdf_path):
            raise RuntimeError("PDF was not created")

        return pdf_path

    except subprocess.TimeoutExpired:
        raise RuntimeError("LibreOffice conversion timed out")
    except Exception as e:
        # Fall back to simple conversion
        return excel_to_pdf_libreoffice_simple(excel_path, output_dir)
    finally:
        if os.path.exists(script_path):
            os.remove(script_path)

def excel_to_pdf_libreoffice_simple(excel_path, output_dir):
    """Simple LibreOffice conversion without scaling (fallback)"""
    libreoffice_path = get_libreoffice_path()

    cmd = [
        libreoffice_path,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', output_dir,
        excel_path
    ]

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=300
    )

    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

    if not os.path.exists(pdf_path):
        raise RuntimeError("PDF was not created by LibreOffice")

    return pdf_path

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

@app.route('/check-libreoffice')
def check_libreoffice():
    """Check if LibreOffice is available for print-style PDF conversion"""
    available = is_libreoffice_available()
    return jsonify({
        'available': available,
        'message': 'LibreOffice is available' if available else 'LibreOffice is not installed'
    })

@app.route('/convert-print', methods=['POST'])
def convert_print():
    """Convert Excel to PDF using LibreOffice (preserves original formatting)"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_excel_file(file.filename):
        return jsonify({'error': 'Invalid file type. Print as PDF only supports .xlsx, .xls, or .xlsm files (not CSV)'}), 400

    if not is_libreoffice_available():
        return jsonify({'error': 'LibreOffice is not available on this server'}), 503

    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{filename}")
        file.save(excel_path)

        # Convert using LibreOffice
        pdf_path = excel_to_pdf_libreoffice(excel_path, app.config['UPLOAD_FOLDER'])

        # Rename PDF to include timestamp
        pdf_filename = filename.rsplit('.', 1)[0] + '.pdf'
        final_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{pdf_filename}")
        os.rename(pdf_path, final_pdf_path)

        # Clean up Excel file
        os.remove(excel_path)

        # Get PDF file size
        pdf_size = os.path.getsize(final_pdf_path)

        return jsonify({
            'success': True,
            'filename': pdf_filename,
            'download_path': f"/download/{timestamp}_{pdf_filename}",
            'metadata': {
                'file_size': pdf_size,
                'method': 'LibreOffice'
            }
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
