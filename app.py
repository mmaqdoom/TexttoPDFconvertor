import sys

# Safety guard: Check if required third-party libraries are installed
try:
    from flask import Flask, render_template, request, send_file, flash, redirect, url_for
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    import docx
    import pandas as pd
except ModuleNotFoundError as e:
    missing_module = e.name
    print("\n" + "=" * 60)
    print(f" ERROR: Required module '{missing_module}' is not installed in the active environment.")
    print(" Please make sure your virtual environment is active and run:")
    print("     pip install -r requirements.txt")
    print("=" * 60 + "\n")
    sys.exit(1)

import os
import io

app = Flask(__name__)
app.secret_key = 'super-secret-key-change-this'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB upload limit

ALLOWED_EXTENSIONS = {'txt', 'docx', 'xlsx', 'xls', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def generate_pdf_from_text(text_content):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    normal_style.fontSize = 10
    normal_style.leading = 13

    story = []
    lines = text_content.splitlines()
    for line in lines:
        if line.strip():
            clean_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(clean_line, normal_style))
        else:
            story.append(Spacer(1, 10))

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_pdf_from_docx(file_stream):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    heading_style = styles['Heading1']

    docx_doc = docx.Document(file_stream)
    story = []

    for para in docx_doc.paragraphs:
        text = para.text.strip()
        if text:
            clean_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            if para.style.name.startswith('Heading'):
                story.append(Paragraph(clean_text, heading_style))
                story.append(Spacer(1, 6))
            else:
                story.append(Paragraph(clean_text, normal_style))
                story.append(Spacer(1, 4))

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_pdf_from_excel(file_stream, filename):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = []

    if filename.endswith('.csv'):
        df = pd.read_csv(file_stream)
        sheets = {'Sheet1': df}
    else:
        sheets = pd.read_excel(file_stream, sheet_name=None)

    for sheet_name, df in sheets.items():
        df = df.fillna('')
        headers = list(df.columns)
        data = [headers] + df.astype(str).values.tolist()

        formatted_data = []
        for row in data:
            formatted_row = [
                Paragraph(str(cell).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'), styles['Normal']) 
                for cell in row
            ]
            formatted_data.append(formatted_row)

        story.append(Paragraph(f"<b>Sheet: {sheet_name}</b>", styles['Heading2']))
        story.append(Spacer(1, 8))

        t = Table(formatted_data, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4A90E2')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F8F9FA')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
        ]))
        story.append(t)
        story.append(Spacer(1, 15))

    doc.build(story)
    buffer.seek(0)
    return buffer

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        raw_text = request.form.get('raw_text', '').strip()
        uploaded_file = request.files.get('file')

        if raw_text:
            pdf_buffer = generate_pdf_from_text(raw_text)
            return send_file(pdf_buffer, as_attachment=True, download_name="converted_text.pdf", mimetype='application/pdf')

        if uploaded_file and uploaded_file.filename != '':
            if not allowed_file(uploaded_file.filename):
                flash('Unsupported file extension. Allowed: .txt, .docx, .xlsx, .xls, .csv')
                return redirect(url_for('index'))

            ext = uploaded_file.filename.rsplit('.', 1)[1].lower()
            file_stream = io.BytesIO(uploaded_file.read())

            if ext == 'txt':
                text_content = file_stream.read().decode('utf-8', errors='ignore')
                pdf_buffer = generate_pdf_from_text(text_content)
            elif ext == 'docx':
                pdf_buffer = generate_pdf_from_docx(file_stream)
            elif ext in ['xlsx', 'xls', 'csv']:
                pdf_buffer = generate_pdf_from_excel(file_stream, uploaded_file.filename)

            download_name = f"{os.path.splitext(uploaded_file.filename)[0]}.pdf"
            return send_file(pdf_buffer, as_attachment=True, download_name=download_name, mimetype='application/pdf')

        flash('Please paste text or select a valid file to convert.')
        return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    print("Starting Flask server locally at http://127.0.0.1:5000")
    app.run(debug=True, port=5000)