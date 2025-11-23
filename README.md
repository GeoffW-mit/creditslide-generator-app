# Flask PPT Generator

This project is a Flask web application that allows users to upload a CSV file and generate PowerPoint slides based on the data. The generated PPT files are saved to a specified Google Drive directory.

## Features
- Upload CSV file with columns: `resource_name`, `attribution_type`, `attribution_text`
- Specify Google Drive template path and output directory
- Generate PPT slides grouped by resource name and attribution type

## Requirements
- Python 3.8+
- Flask
- pandas
- python-pptx

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/my-flask-ppt-app.git
   cd my-flask-ppt-app
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. Run the Flask app:
   ```bash
   python app.py
   ```

2. Open your browser and go to:
   ```
   http://127.0.0.1:5000/
   ```

3. Upload your CSV file, enter the template path and output directory, then click **Generate PPT**.

## Project Structure
```
my-flask-ppt-app/
│
├── app.py                # Flask application
├── generate_slides.py    # PPT generation function
├── requirements.txt       # Dependencies
├── templates/            # HTML templates
│   └── index.html         # Upload form
├── uploads/              # Temporary CSV storage
└── README.md             # Project documentation
```

## Deployment
You can deploy this app on platforms like **Render**, **Heroku**, or **PythonAnywhere**. For deployment:
- Ensure `requirements.txt` is included.
- Configure environment variables for Google Drive API credentials if needed.

## License
MIT License
