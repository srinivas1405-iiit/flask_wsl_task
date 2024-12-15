from flask import Flask, request, jsonify, render_template
import pandas as pd
import os
import tempfile

app = Flask(__name__)

# Defining the validation function
def validate_excel(file_path):
    required_sheets = {
        "Course": ["Course ID", "Course Name"],
        "Topic": ["Topic ID", "Topic Name", "Description"],
        "Resource": ["Resource ID", "Resource Name", "Resource Content", "Module ID", "Module Name", "Sub Module ID"],
        "Learner": ["Learner ID", "Name", "Essay", "Module ID", "Submodule ID"]
    }
    
    try:
        # Loading the excel file
        with pd.ExcelFile(file_path) as excel_data:
            validation_results = {
                "is_valid": True,
                "errors": []
            }
            
            # Checking for required excel sheets in excel file
            for sheet_name, required_columns in required_sheets.items():
                if sheet_name not in excel_data.sheet_names:
                    validation_results["is_valid"] = False
                    validation_results["errors"].append(f"Missing sheet: {sheet_name}")
                else:
                    sheet_df = pd.read_excel(file_path, sheet_name=sheet_name)
                    
                    # Checking for required columns
                    missing_columns = [col for col in required_columns if col not in sheet_df.columns]
                    if missing_columns:
                        validation_results["is_valid"] = False
                        validation_results["errors"].append(f"Missing columns in {sheet_name}: {', '.join(missing_columns)}")
                    
                    # Checking for non-empty sheets
                    if sheet_df.empty:
                        validation_results["is_valid"] = False
                        validation_results["errors"].append(f"Sheet {sheet_name} is empty")
        
        return validation_results
    
    except Exception as e:
        return {
            "is_valid": False,
            "errors": [str(e)]
        }

# Defining the route for file upload and validation
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('upload.html', validation_results={"is_valid": False, "errors": ["No file part"]})
        
        file = request.files['file']
        if file.filename == '':
            return render_template('upload.html', validation_results={"is_valid": False, "errors": ["No selected file"]})
        
        # Save the file to a temporary directory
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            file_path = tmp_file.name
            file.save(file_path)
        
        # Validating the excel file
        validation_results = validate_excel(file_path)
        
        # Remove the temporary file after validation
        os.remove(file_path)
        
        return render_template('upload.html', validation_results=validation_results)
    
    return render_template('upload.html', validation_results=None)

if __name__ == '__main__':
    app.run(debug=True)
