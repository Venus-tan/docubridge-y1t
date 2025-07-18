from flask import Flask, request, render_template
import os
import pandas as pd

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Create uploads folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    print("\n=== DEBUG: Upload Check ===")
    print("Files received:", request.files)
    print("Form data:", request.form)

    if 'excel_file' not in request.files:
        return "No file uploaded", 400

    excel_file = request.files['excel_file']
    user_question = request.form.get('Question', '')

    if excel_file.filename == '':
        return "No selected file", 400

    # Save the uploaded file
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(filepath)

    try:
        # Process the Excel file
        df = pd.read_excel(filepath)
        preview = df.head(3).to_html(index=False)
        
        # Clean up
        os.remove(filepath)
        
        return render_template('index.html', preview=preview)
    
    except Exception as e:
        return f"Error processing file: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)