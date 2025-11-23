
from flask import Flask, request, render_template, send_from_directory
import pandas as pd
import os
import zipfile
from slide_generator import create_ppt_from_dataframe

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    csv_file = request.files['csv_file']
    template_path = request.form['template_path']

    # ✅ Clear old files before generating new ones
    for old_file in os.listdir(OUTPUT_FOLDER):
        os.remove(os.path.join(OUTPUT_FOLDER, old_file))

    # Save CSV temporarily
    csv_path = os.path.join(UPLOAD_FOLDER, csv_file.filename)
    csv_file.save(csv_path)

    # Generate PPT files locally
    df = pd.read_csv(csv_path)
    create_ppt_from_dataframe(df, template_path, OUTPUT_FOLDER)

    # Create ZIP file of all PPTs
    zip_filename = "generated_slides.zip"
    zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for filename in os.listdir(OUTPUT_FOLDER):
            if filename.endswith(".pptx"):
                zipf.write(os.path.join(OUTPUT_FOLDER, filename), filename)

    return f"""
    <h3>✅ PPT files generated!</h3>
    <p>/download_zip/{zip_filename}Download All as ZIP</a></p>
    """

@app.route('/download_zip/<filename>')
def download_zip(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
