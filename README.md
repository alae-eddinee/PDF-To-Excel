# PDF to Excel Pivot Converter

A Streamlit web application that converts BON DE COMMANDE PDF files to Excel pivot tables.

## Features

- 📄 **PDF Upload**: Drag and drop or select multiple PDF files
- 🔄 **Automatic Format Detection**: Supports Marjane and LV formats
- 📊 **Excel Generation**: Creates pivot tables with article and store data
- 📥 **Download**: Download generated Excel files instantly
- 🎨 **Modern UI**: Clean, user-friendly interface

## Supported Formats

### Marjane Format
- BON DE COMMANDE — MEDIDIS / MARJANE HOLDING
- Extracts EAN, product names, and quantities by store

### LV Format  
- BON DE COMMANDE — MEDIDIS / LV
- Supports HYPER MARCHE LV and HYPER SUD variants
- Handles multi-line product descriptions with dimensions

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the Streamlit app:
```bash
streamlit run app.py
```

The app will open automatically in your browser at `http://localhost:8501`

## Usage

1. **Upload PDF Files**: Click "Browse files" or drag and drop PDF files
2. **Process**: Click the "Generate Excel Files" button
3. **Download**: Download individual Excel files for each processed PDF
4. **View Results**: Check detailed statistics for each file

## Command Line Usage

You can also use the original script directly:

```bash
# Process all PDFs in current directory
python bc_pdf_to_pivot.py

# Process specific PDF
python bc_pdf_to_pivot.py file.pdf

# Process with custom output name
python bc_pdf_to_pivot.py file.pdf output.xlsx
```

## File Structure

```
├── app.py                 # Streamlit web application
├── bc_pdf_to_pivot.py     # Core PDF processing logic
├── requirements.txt       # Python dependencies
├── README.md             # This file
└── *.pdf                 # Your PDF files
```

## Dependencies

- **streamlit**: Web application framework
- **pdfplumber**: PDF text extraction
- **pandas**: Data manipulation
- **openpyxl**: Excel file generation
- **Pillow**: Image processing

## Troubleshooting

- **PDF not processing**: Ensure the PDF contains recognizable Marjane or LV format headers
- **Empty Excel file**: Check if the PDF has valid article data with EAN codes
- **Format detection fails**: The app defaults to Marjane format if unsure

## Technical Details

The app processes PDFs by:
1. Detecting format based on text content
2. Extracting article data (EAN, description, quantities)
3. Identifying store locations
4. Building pivot tables with totals
5. Generating formatted Excel files with styling
