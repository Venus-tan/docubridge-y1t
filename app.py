from flask import Flask, request, render_template, redirect, url_for
import os
import pandas as pd

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xls', 'xlsx'}

# Create uploads folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if the post request has the file part
    if 'excel_file' not in request.files:
        return "No file part in the request", 400
    
    excel_file = request.files['excel_file']
    user_question = request.form.get('user_question', '').strip()
    
    # If user does not select file, browser submits an empty file without filename
    if excel_file.filename == '':
        return "No selected file", 400
    
    if not allowed_file(excel_file.filename):
        return "Invalid file type. Only Excel files (.xls, .xlsx) are allowed.", 400
    
    # Save the uploaded file temporarily
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(filepath)
    
    try:
        # Day 4 Task: Read and parse Excel content
        # Get sheet names first
        excel_data = pd.ExcelFile(filepath)
        sheet_names = excel_data.sheet_names
        
        # Read the first sheet by default (can be expanded to handle multiple sheets)
        df = pd.read_excel(filepath, sheet_name=sheet_names[0])
        
        # Get some basic info about the data
        num_rows, num_cols = df.shape
        column_names = list(df.columns)
        
        # Day 3 Task: Simple confirmation response with file info and question
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
        
        # Render a response page with the confirmation and data preview
        return render_template('response.html', data=response_data)
    
    except Exception as e:
        # Clean up in case of error
        if os.path.exists(filepath):
            os.remove(filepath)
        return f"Error processing the file: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)