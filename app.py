import os
from flask import Flask, render_template, request, send_file, session, redirect, jsonify
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from idle_time_analysis import perform_idle_time_analysis
import tempfile
from datetime import time

CSV_HEADER_ROW = 2

app = Flask(__name__, static_url_path='/static')
app.secret_key = 'your_secret_key'  # Add a secret key for sessions

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'csv'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Define the allowed_file function
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# Initialize df in the session
def initialize_session_data():
    session['df'] = None

@app.route('/')
def home():
    # Initialize session data
    initialize_session_data()
    return render_template('upload.html')


# 1. Update Flask Routes



    
@app.route('/idle_time_analysis', methods=['GET', 'POST'])
def idle_time_analysis():
    if request.method == 'POST':
        # Handle POST request
        username = request.form['username']
        # Perform the idle time analysis
        df = pd.read_csv(session['df'], skiprows=3)  # Read CSV while skipping the first 3 rows
        df['DateTime'] = pd.to_datetime(df['DateTime'], format='%I:%M %p')

        # Define the start and end times as time objects
        start_time = pd.Timestamp('1900-01-01 20:00:00').time()  # 8:00 PM
        end_time = pd.Timestamp('1900-01-01 23:59:59').time()    # 11:59:59 PM

        # Extract the time part from DateTime
        df['Time'] = df['DateTime'].dt.time

        # Filter data between start_time and end_time
        user_data = df[(df['UserID'] == username) & (df['Time'] >= start_time) & (df['Time'] <= end_time)]

        # Calculate TimeDiff within user_data
        user_data['TimeDiff'] = user_data['DateTime'].diff().dt.total_seconds() / 60
        # Make sure to perform this calculation only if user_data is not empty to avoid errors.

        return render_template('user_data.html', username=username, user_data=user_data, start_time=start_time, end_time=end_time)
    elif request.method == 'GET':
        # Handle GET request, show a form to input username
        return render_template('idle_time_input.html')








@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    if file and allowed_file(file.filename):
        # Create a temporary file to store the uploaded file
        temp_file = tempfile.NamedTemporaryFile(delete=False)
        filename = temp_file.name
        file.save(filename)

        # Process the uploaded CSV file
        df = pd.read_csv(filename, header=CSV_HEADER_ROW)
        df['DateTime'] = pd.to_datetime(df['DateTime'], format='%I:%M %p')

       
        # Store the file path in the session
        session['df'] = filename

        # Create a new Excel workbook
        output_excel_file = os.path.join(app.config['OUTPUT_FOLDER'], 'processed_data.xlsx')
        book = openpyxl.Workbook()

        # Perform analysis for different pick types and add the results to the Excel workbook
        from putwall_pick import perform_putwall_pick_analysis
        perform_putwall_pick_analysis(df, book)

        from single_packing import perform_single_pack_analysis
        perform_single_pack_analysis(df, book)

        from multi_pack import perform_multi_pack_analysis
        perform_multi_pack_analysis(df, book)

        from regular_pick import perform_regular_pick_analysis
        perform_regular_pick_analysis(df, book)

        from single_pick import perform_single_pick_analysis
        perform_single_pick_analysis(df, book)

        

        from resolve import peform_resolve_analysis
        peform_resolve_analysis(df, book)   

        from replenishment_pick import perform_replenishment_pick_analysis
        perform_replenishment_pick_analysis(df, book)

        from quick_move import peform_quick_move_analysis
        peform_quick_move_analysis(df, book)

        from idle_time import perform_idle_time_analysis
        perform_idle_time_analysis(df, book)

        from hourly_pick_totals import perform_hourly_pick_totals_analysis
        perform_hourly_pick_totals_analysis(df, book)

        from hourly_packs_total import perform_hourly_pack_totals_analysis
        perform_hourly_pack_totals_analysis(df, book)

        from singleunitspickedbyzone import perform_single_pick_byzone
        perform_single_pick_byzone(df, book)



        # from picks_per_zones import perform_pick_totals_analysis_per_zones
        # perform_pick_totals_analysis_per_zones(df, book)


        

        # Save the Excel file
        book.save(output_excel_file)

        # Provide a link to download the processed file
        download_link = f'<a href="/download/{output_excel_file}">Download Processed Excel File</a>'
        return redirect('/processing_done') 
    




@app.route('/processing_done')
def processing_done():
    # You can customize this route as needed, and render a template or add additional logic
    return render_template('processing_done.html')

@app.route('/to_do_list')
def todolist():
    return render_template('to_do_list.html')

@app.route('/download/<path:filename>')
def download_file(filename):
    print(f"Downloading file: {filename}")  # Add this line for debugging
    return send_file(filename, as_attachment=True)

@app.route('/get_user_data', methods=['POST'])
def get_user_data():
    user_ids = request.form['user_ids'].split(',')
    user_ids = [uid.strip() for uid in user_ids]

    try:
        excel_file = os.path.join(app.config['OUTPUT_FOLDER'], 'processed_data.xlsx')
        sheet_names = ['REPLENISHMENT PICK', 'SINGLE PICK', 'REGULAR PICK', 'PUTWALL PICKING']
        tables = {}

        for sheet_name in sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            if not df.empty and df['UserID'].isin(user_ids).any():
                filtered_df = df[df['UserID'].isin(user_ids)]
                tables[sheet_name] = filtered_df.to_html(classes='data', index=False)

        if tables:
            return render_template('display_user_data.html', tables=tables)
        else:
            return render_template('no_data_found.html')  # You should create this template

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    app.run(host='0.0.0.0',port=5000)
