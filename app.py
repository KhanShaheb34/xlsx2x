from flask import Flask, request, send_file
from xlsx2x import generatePDF, generatePNG

app = Flask(__name__)

@app.route('/png', methods=['POST'])
def get_png():
    XLSXFile = request.files['xlsx'].read()
    with open('sheet.xlsx', 'wb') as f:
        f.write(XLSXFile)
    generatePNG('sheet.xlsx', 'out.png')
    return send_file('out.png', mimetype='image/png')

@app.route('/pdf', methods=['POST'])
def get_pdf():
    XLSXFile = request.files['xlsx'].read()
    with open('sheet.xlsx', 'wb') as f:
        f.write(XLSXFile)
    generatePDF('sheet.xlsx', 'out.pdf')
    return send_file('out.pdf', mimetype='document/pdf')