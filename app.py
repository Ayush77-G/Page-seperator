from flask import Flask, render_template, request, jsonify
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

app = Flask(__name__)

def save_pages_separately(pdf_file, custom_names=None, output_dir=None):
    if output_dir is None:
        output_dir = ''
    elif not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    with open(pdf_file, 'rb') as file:
        reader = PdfReader(file)
        num_pages = len(reader.pages)
        
        if custom_names is None:
            custom_names = [f'page_{i+1}.pdf' for i in range(num_pages)]
        elif len(custom_names) != num_pages:
            raise ValueError("Number of custom names does not match the number of pages.")
        
        for i, (page, custom_name) in enumerate(zip(reader.pages, custom_names)):
            writer = PdfWriter()
            writer.add_page(page)
            
            output_path = os.path.join(output_dir, custom_name)
            if not output_path.lower().endswith('.pdf'):
                output_path += '.pdf'
                
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
                print(f'Page {i+1} saved as {output_path}.')

def read_from_excel(excel_file):
    df = pd.read_excel(excel_file)
    custom_names = df['CustomName'].tolist()
    return custom_names

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/extract', methods=['POST'])
def extract_pages():
    doc_file = request.files['docFile']
    excel_file = request.files['excelFile']
    output_dir = request.form['outputDir']

    if not doc_file or not excel_file or not output_dir:
        return jsonify({'success': False, 'message': 'Please provide all required input.'}), 400

    doc_file_path = os.path.join('uploads', doc_file.filename)
    excel_file_path = os.path.join('uploads', excel_file.filename)
    doc_file.save(doc_file_path)
    excel_file.save(excel_file_path)

    custom_names = read_from_excel(excel_file_path)
    save_pages_separately(doc_file_path, custom_names, output_dir=output_dir)

    return jsonify({'success': True, 'message': 'Pages extracted successfully.'})

if __name__ == '__main__':
    app.run(debug=True)
