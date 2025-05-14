# Excel VLOOKUP Generator

A beautiful, user-friendly Streamlit app that automatically generates VLOOKUP formulas between Excel files.

## Features

- Modern, intuitive user interface with animations
- Supports multiple Excel sheet formats
- Preview of processed data
- Detailed file information
- Progress tracking
- Error handling with IFERROR formulas
- Responsive design

## Getting Started

### Prerequisites

- Python 3.7+
- pip

### Installation

1. Clone this repository or download the files

2. Install the required packages:

```bash
pip install -r requirements.txt
```

### Running the app locally

```bash
streamlit run app.py
```

This will start the Streamlit server and open the app in your default web browser at http://localhost:8501.

## How to Use

1. Upload your main Excel file where you want to add VLOOKUP formulas
2. Upload your lookup Excel file (O.xlsx)
3. Select the sheet to process
4. Enter the last row to process
5. Click "Generate VLOOKUP Formulas"
6. Download the processed file

## Project Structure

- `app.py` - Main application file with UI components
- `vlookup_processor.py` - Backend logic for processing Excel files
- `requirements.txt` - Required Python packages

## Customization

You can customize the appearance by modifying the CSS in the `local_css()` function in `app.py`.

## Troubleshooting

- If you encounter issues with file uploads, check if your Excel files are properly formatted and not corrupted
- For large files, processing may take longer - the progress bar will help track the status

## License

This project is open source.
