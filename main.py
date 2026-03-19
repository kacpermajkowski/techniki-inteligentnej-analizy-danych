from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK

from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

from reportlab.pdfgen import canvas

from flask import Flask, render_template, request
import pandas
import os.path

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html', output="test")

@app.route('/upload-file', methods=['POST'])
def upload_file():
    file = request.files['file']
    if not file or file.filename == '':
        return "No file uploaded", 400
    if not file.filename.endswith('.xlsx'):
        return "Invalid file type. Supported types: .xlsx", 400

    df = pandas.read_excel(file)
    mode = request.form['mode']
    if mode == 'pdf':
        convert_to_pdf(file.filename, df)
    elif mode == 'docx':
        convert_to_docx(file.filename, df)
    elif mode == 'both':
        convert_to_pdf(file.filename, df)
        convert_to_docx(file.filename, df)
    else:
        return "Invalid mode. Supported modes: pdf, docx, both", 400

    return 'File processed successfully', 200

def convert_to_pdf(filename, df):
    pass

def convert_to_docx(filename, df):
    document = Document()

    for i, row in df.iterrows():
        hp = document.add_paragraph(filename + ' | Row ' + str(i + 1) + ' of ' + str(len(df)) + '\n')
        insertHR(hp)
        last_p = None
        for col in df.columns:
            p = document.add_paragraph()
            p.add_run(col + ': ').bold = True
            p.add_run(str(row[col]))
            last_p = p
        last_p.add_run().add_break(WD_BREAK.PAGE)

    document.add_heading('Test title', 0)
    document.add_paragraph(' ')

    if os.path.isfile(filename):
        os.remove(filename + '.docx')

    document.save(filename + '.docx')

# Source: https://github.com/python-openxml/python-docx/issues/105#issuecomment-442786431

def insertHR(paragraph):
    p = paragraph._p  # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)


if __name__ == '__main__':
    app.run(debug=True)


