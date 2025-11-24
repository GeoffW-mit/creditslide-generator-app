
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
        print("\n=== ✅ /process route hit ===")

        csv_file = request.files.get('csv_file')
        template_file = request.files.get('template_file')
        if not csv_file or not template_file:
            print("❌ Missing files in request.")
            return jsonify({"success": False, "error": "CSV or template file missing."})

        print(f"Uploaded files: CSV={csv_file.filename}, Template={template_file.filename}")

        # Clear old files
        for old_file in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, old_file))
        print("✅ Old files cleared from output folder.")

        # Save uploaded files
        csv_path = os.path.join(UPLOAD_FOLDER, csv_file.filename)
        template_path = os.path.join(UPLOAD_FOLDER, template_file.filename)
        csv_file.save(csv_path)
        template_file.save(template_path)
        print(f"✅ Files saved: CSV={csv_path}, Template={template_path}")

        # Load CSV
        df = pd.read_csv(csv_path)
        print(f"✅ CSV loaded successfully with {len(df)} rows and {len(df.columns)} columns.")

        # Generate PPT files
        print("▶ Calling create_ppt_from_dataframe()...")
        create_ppt_from_dataframe(df, template_path, OUTPUT_FOLDER)
        print("✅ PPT generation function executed.")

        # Check output folder contents
        output_files = os.listdir(OUTPUT_FOLDER)
        print("✅ Output folder contents after generation:", output_files)

        # Verify PPT files exist
        ppt_files = [f for f in output_files if f.endswith('.pptx')]
        if not ppt_files:
            print("❌ No PPT files generated.")
            return jsonify({"success": False, "error": "No PPT files were generated."})

        # Create ZIP
        zip_filename = "generated_slides.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for filename in ppt_files:
                zipf.write(os.path.join(OUTPUT_FOLDER, filename), filename)
        print(f"✅ ZIP file created at {zip_path}")

        return jsonify({
            "success": True,
            "download_url": f"/download_zip/{zip_filename}"
        })

    except Exception as e:
        print("❌ Error occurred:", str(e))
        return jsonify({"success": False, "error": str(e)})

@app.route('/download_zip/<filename>')
def download_zip(filename):
    print(f"▶ Download request for {filename}")
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    print("🚀 Flask app starting...")
    app.run(debug=True)
