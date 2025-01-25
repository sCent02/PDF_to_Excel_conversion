from flask import Flask, request, send_file
import pandas as pd
import pdfplumber

app = Flask(__name__)

@app.route('/')
def home():
    return '<h1>PDF to Excel Converter</h1><form action="/convert" method="post" enctype="multipart/form-data"><input type="file" name="file"><button type="submit">Convert</button></form>'

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file uploaded", 400
    file = request.files['file']
    if file.filename == '':
        return "No file selected", 400

    try:
        with pdfplumber.open(file) as pdf:
            data = []
            for page in pdf.pages:
                data.extend(page.extract_table())

            df = pd.DataFrame(data)
            output_file = "output.xlsx"
            df.to_excel(output_file, index=False)

            return send_file(output_file, as_attachment=True)
    except Exception as e:
        return str(e), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)