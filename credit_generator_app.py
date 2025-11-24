
from flask import Flask, request, render_template, send_from_directory, jsonify
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
    try:
        csv_file = request.files['csv_file']
        template_file = request.files['template_file']

        # Clear old files
        for old_file in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, old_file))

        # Save uploaded files
        csv_path = os.path.join(UPLOAD_FOLDER, csv_file.filename)
        template_path = os.path.join(UPLOAD_FOLDER, template_file.filename)
        csv_file.save(csv_path)
        template_file.save(template_path)

        # Generate PPT files
        df = pd.read_csv(csv_path)
        create_ppt_from_dataframe(df, template_path, OUTPUT_FOLDER)

        # Create ZIP
        zip_filename = "generated_slides.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for filename in os.listdir(OUTPUT_FOLDER):
                if filename.endswith('.pptx'):
                    zipf.write(os.path.join(OUTPUT_FOLDER, filename), filename)

        return jsonify({
            "success": True,
            "download_url": f"/download_zip/{zip_filename}"
        })

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

@app.route('/download_zip/<filename>')
def download_zip(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
