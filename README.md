# DOCX Formatter - AI Tool

An AI-powered web application that processes DOCX files to automatically format headings and arrange content in a 2-column layout.

## Features

- **Bold Headings**: Automatically detects and makes headings bold and centered
- **2-Column Layout**: Arranges document content in a clean 2-column format
- **Smart Detection**: Uses AI logic to identify headings vs content
- **Web Interface**: Easy-to-use drag-and-drop web interface
- **File Processing**: Handles DOCX files up to 16MB

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python app.py
```

3. Open your browser and go to `http://localhost:5000`

## Usage

1. Open the web application in your browser
2. Upload a DOCX file by clicking or dragging and dropping
3. Click "Process Document" 
4. Download the formatted document

## How it Works

The tool analyzes your DOCX file and:
- Identifies headings (ALL CAPS text, short phrases, or text ending with ':')
- Makes headings bold and centers them
- Splits remaining content into two balanced columns
- Maintains proper spacing and formatting
- Preserves document structure

## File Structure

```
docx-formatter/
├── app.py              # Main Flask application
├── templates/
│   └── index.html      # Web interface
├── uploads/            # Temporary file storage
├── requirements.txt    # Python dependencies
└── README.md          # This file
```

## Requirements

- Python 3.7+
- Flask
- python-docx
- Modern web browser