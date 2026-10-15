from flask import (
    Flask,
    jsonify,
    render_template,
    send_from_directory,
    url_for,
)
from werkzeug.utils import secure_filename

import os
import shutil
import uuid
import zipfile

import pandas as pd

from slide_generator import create_ppt_from_dataframe


app = Flask(__name__)

BASE_DIR = app.root_path
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

# Upload size limit for the CSV, template, and font upload(s) combined.
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

# Limit extracted font files from one ZIP to reduce risk from oversized archives.
MAX_EXTRACTED_FONT_BYTES = 50 * 1024 * 1024

os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


def save_uploaded_fonts(font_uploads, fonts_dir):
    """
    Save individual .ttf/.otf uploads and extract supported files from ZIPs.

    ZIP contents are copied by filename into fonts_dir; archive paths are not
    extracted directly, preventing paths in the ZIP from writing elsewhere.
    """
    os.makedirs(fonts_dir, exist_ok=True)
    total_extracted_bytes = 0
    font_count = 0

    for uploaded_file in font_uploads:
        original_name = uploaded_file.filename or ""
        extension = os.path.splitext(original_name)[1].lower()

        if extension == ".zip":
            uploaded_file.stream.seek(0)

            try:
                archive = zipfile.ZipFile(uploaded_file.stream)
            except zipfile.BadZipFile as exc:
                raise ValueError(
                    f"{original_name!r} is not a valid ZIP file."
                ) from exc

            with archive:
                for info in archive.infolist():
                    if info.is_dir():
                        continue

                    normalized_name = info.filename.replace("\\", "/")
                    basename = os.path.basename(normalized_name)
                    file_extension = os.path.splitext(basename)[1].lower()

                    if file_extension not in {".ttf", ".otf"}:
                        continue

                    total_extracted_bytes += info.file_size
                    if total_extracted_bytes > MAX_EXTRACTED_FONT_BYTES:
                        raise ValueError(
                            "The extracted font files exceed the allowed size."
                        )

                    safe_name = secure_filename(basename)
                    if not safe_name:
                        continue

                    destination = os.path.join(fonts_dir, safe_name)

                    with archive.open(info) as source, open(destination, "wb") as target:
                        shutil.copyfileobj(source, target)

                    font_count += 1

        elif extension in {".ttf", ".otf"}:
            safe_name = secure_filename(original_name)
            if not safe_name:
                raise ValueError(
                    f"Could not use the uploaded font filename "
                    f"{original_name!r}."
                )

            destination = os.path.join(fonts_dir, safe_name)
            uploaded_file.save(destination)

            total_extracted_bytes += os.path.getsize(destination)
            if total_extracted_bytes > MAX_EXTRACTED_FONT_BYTES:
                raise ValueError(
                    "The uploaded font files exceed the allowed size."
                )

            font_count += 1

        else:
            raise ValueError(
                f"Unsupported font upload {original_name!r}. "
                "Upload a ZIP file or .ttf/.otf font files."
            )

    if font_count == 0:
        raise ValueError(
            "No .ttf or .otf files were found. Upload the font ZIP or "
            "the individual font files required by the template."
        )

    return font_count


@app.errorhandler(413)
def request_too_large(_error):
    return jsonify({
        "success": False,
        "error": "The combined upload exceeds the 100 MB limit.",
    }), 413


@app.route("/process", methods=["POST"])
def process():
    job_id = uuid.uuid4().hex
    job_dir = os.path.join(OUTPUT_FOLDER, job_id)

    try:
        csv_file = request.files.get("csv_file")
        template_file = request.files.get("template_file")
        font_uploads = [
            item
            for item in request.files.getlist("font_files")
            if item and item.filename
        ]

        if not csv_file or not csv_file.filename:
            return jsonify({
                "success": False,
                "error": "Please upload a CSV file.",
            }), 400

        if not template_file or not template_file.filename:
            return jsonify({
                "success": False,
                "error": "Please upload a PowerPoint template.",
            }), 400

        if not font_uploads:
            return jsonify({
                "success": False,
                "error": "Please upload the font ZIP or individual .ttf/.otf files.",
            }), 400

        csv_name = secure_filename(csv_file.filename)
        template_name = secure_filename(template_file.filename)

        if not csv_name.lower().endswith(".csv"):
            return jsonify({
                "success": False,
                "error": "The attribution file must be a .csv file.",
            }), 400

        if not template_name.lower().endswith(".pptx"):
            return jsonify({
                "success": False,
                "error": "The template must be a .pptx file.",
            }), 400

        work_dir = os.path.join(job_dir, "work")
        slides_dir = os.path.join(job_dir, "slides")
        fonts_dir = os.path.join(work_dir, "fonts")

        os.makedirs(work_dir, exist_ok=True)
        os.makedirs(slides_dir, exist_ok=True)

        csv_path = os.path.join(work_dir, csv_name)
        template_path = os.path.join(work_dir, template_name)

        csv_file.save(csv_path)
        template_file.save(template_path)

        font_count = save_uploaded_fonts(font_uploads, fonts_dir)

        df = pd.read_csv(csv_path)

        required_columns = {
            "attribution_type",
            "attribution_text",
            "resource_name",
        }
        missing_columns = required_columns.difference(df.columns)

        if missing_columns:
            raise ValueError(
                "The CSV is missing required column(s): "
                + ", ".join(sorted(missing_columns))
            )

        create_ppt_from_dataframe(
            df=df,
            template_path=template_path,
            output_dir=slides_dir,
            font_dir=fonts_dir,
        )

        ppt_files = [
            filename
            for filename in os.listdir(slides_dir)
            if filename.lower().endswith(".pptx")
        ]

        if not ppt_files:
            raise RuntimeError("No PowerPoint files were generated.")

        zip_path = os.path.join(job_dir, "generated_slides.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for filename in ppt_files:
                file_path = os.path.join(slides_dir, filename)
                zip_file.write(file_path, arcname=filename)

        # Delete the uploaded CSV, template, and font files after generation.
        shutil.rmtree(work_dir, ignore_errors=True)
        shutil.rmtree(slides_dir, ignore_errors=True)

        return jsonify({
            "success": True,
            "font_files_used": font_count,
            "download_url": url_for("download_zip", job_id=job_id),
        })

    except Exception as exc:
        app.logger.exception("Slide generation failed")

        # Delete this job's files if processing failed.
        shutil.rmtree(job_dir, ignore_errors=True)

        return jsonify({
            "success": False,
            "error": str(exc),
        }), 400


@app.route("/download_zip/<job_id>")
def download_zip(job_id):
    # Only accept IDs in the format generated by uuid.uuid4().hex.
    if len(job_id) != 32 or any(
        character not in "0123456789abcdef"
        for character in job_id.lower()
    ):
        return jsonify({
            "success": False,
            "error": "Invalid download ID.",
        }), 404

    job_dir = os.path.join(OUTPUT_FOLDER, job_id)
    zip_path = os.path.join(job_dir, "generated_slides.zip")

    if not os.path.isfile(zip_path):
        return jsonify({
            "success": False,
            "error": "The ZIP file was not found. Please generate the slides again.",
        }), 404

    response = send_from_directory(
        job_dir,
        "generated_slides.zip",
        as_attachment=True,
        download_name="generated_slides.zip",
    )

    # Remove the generated ZIP after the download response is closed.
    response.call_on_close(
        lambda: shutil.rmtree(job_dir, ignore_errors=True)
    )

    return response


if __name__ == "__main__":
    # Do not use debug=True for the deployed Render service.
    app.run()
