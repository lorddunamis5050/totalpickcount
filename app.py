import os
from flask import Flask, render_template, request, send_file
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

CSV_HEADER_ROW = 2 

app = Flask(__name__, static_url_path='/static')



UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'csv'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    if file and allowed_file(file.filename):
        # Save the uploaded file
        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)

        # Process the uploaded CSV file
        df = pd.read_csv(filename, header=CSV_HEADER_ROW)
        df['DateTime'] = pd.to_datetime(df['DateTime'], format='%I:%M %p')

        # Create a new Excel workbook
        output_excel_file = os.path.join(app.config['OUTPUT_FOLDER'], 'processed_data.xlsx')
        book = openpyxl.Workbook()

        # Perform analysis for different pick types and add the results to the Excel workbook
        from putwall_pick import perform_putwall_pick_analysis
        perform_putwall_pick_analysis(df, book)

        from regular_pick import perform_regular_pick_analysis
        perform_regular_pick_analysis(df, book)

        from single_pick import perform_single_pick_analysis
        perform_single_pick_analysis(df, book)

        from replenishment_pick import perform_replenishment_pick_analysis
        perform_replenishment_pick_analysis(df, book)

        from quick_move import peform_quick_move_analysis
        peform_quick_move_analysis(df, book)

        from idle_time import perform_idle_time_analysis
        perform_idle_time_analysis(df, book)

        from hourly_pick_totals import perform_hourly_pick_totals_analysis
        perform_hourly_pick_totals_analysis(df, book)

        # Save the Excel file
        book.save(output_excel_file)

        # Provide a link to download the processed file
        download_link = f'<a href="/download/{output_excel_file}">Download Processed Excel File</a>'

        return render_template('processing_done.html', download_link=download_link)

@app.route('/download/<path:filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    app.run(debug=True)
