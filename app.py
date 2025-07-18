from flask import Flask, request, render_template
import os
import pandas as pd
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xls', 'xlsx'}

# Create uploads folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def log_to_terminal(message):
    """Helper function to format and print messages to terminal"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

@app.route('/')
def home():
    log_to_terminal("Home page accessed")
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Log the start of the upload process
    log_to_terminal("File upload request received")
    
    # Check if the post request has the file part
    if 'excel_file' not in request.files:
        error_msg = "No file part in the request"
        log_to_terminal(error_msg)
        return error_msg, 400
    
    excel_file = request.files['excel_file']
    user_question = request.form.get('user_question', '').strip()
    
    # Log the received file and question
    log_to_terminal(f"Received file: {excel_file.filename} (Size: {len(excel_file.read())} bytes)")
    excel_file.seek(0)  # Reset file pointer after reading size
    
    if user_question:
        log_to_terminal(f"User question: '{user_question}'")
    else:
        log_to_terminal("No question provided by user")
    
    # If user does not select file, browser submits an empty file without filename
    if excel_file.filename == '':
        error_msg = "No selected file"
        log_to_terminal(error_msg)
        return error_msg, 400
    
    if not allowed_file(excel_file.filename):
        error_msg = f"Invalid file type: {excel_file.filename}"
        log_to_terminal(error_msg)
        return "Invalid file type. Only Excel files (.xls, .xlsx) are allowed.", 400
    
    # Save the uploaded file temporarily
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(filepath)
    log_to_terminal(f"File saved temporarily at: {filepath}")
    
    try:
        # Read and parse Excel content
        log_to_terminal("Reading Excel file...")
        
        # Get sheet names first
        excel_data = pd.ExcelFile(filepath)
        sheet_names = excel_data.sheet_names
        log_to_terminal(f"Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")
        
        # Read the first sheet by default
        df = pd.read_excel(filepath, sheet_name=sheet_names[0])
        log_to_terminal(f"Successfully read sheet: '{sheet_names[0]}'")
        
        # Get some basic info about the data
        num_rows, num_cols = df.shape
        column_names = list(df.columns)
        
        # Log data statistics without showing full table
        log_to_terminal(f"Data dimensions: {num_rows} rows × {num_cols} columns")
        log_to_terminal(f"Columns detected: {', '.join(column_names)}")
        
        # Prepare response data
        response_data = {
            'filename': excel_file.filename,
            'user_question': user_question,
            'sheet_name': sheet_names[0],
            'num_rows': num_rows,
            'num_cols': num_cols,
            'column_names': column_names,
            'preview_data': df.head(3).to_html(index=False, classes='table table-striped')
        }
        
        # Clean up - remove the temporary file
        os.remove(filepath)
        log_to_terminal("Temporary file cleaned up")
        
        # Log successful processing
        log_to_terminal("File processed successfully. Preparing response...")
        
        # Render a response page with the confirmation and data preview
        return render_template('response.html', data=response_data)
    
    except Exception as e:
        # Log the error
        error_msg = f"Error processing file: {str(e)}"
        log_to_terminal(error_msg)
        
        # Clean up in case of error
        if os.path.exists(filepath):
            os.remove(filepath)
            log_to_terminal("Cleaned up temporary file after error")
        
        return error_msg, 500

if __name__ == '__main__':
    log_to_terminal("Starting Flask application...")
    app.run(debug=True)