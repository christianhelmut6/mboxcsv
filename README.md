# DCI Mbox Converter

A powerful web application for converting mbox (mailbox) files to Excel, CSV, and TXT formats with no conversion limits.

## Features

- ✅ **No Conversion Limits** - Convert unlimited number of emails
- 📊 **Excel Export** - Auto-formatted columns with proper sizing
- 📄 **CSV Export** - Perfect for data analysis and processing
- 📝 **TXT Export** - Human-readable format for easy review
- 🔧 **Smart Processing** - Handles MIME encoding and various character sets
- 🛡️ **Secure** - All processing is done locally, no data sent to external servers
- 🚀 **Fast** - Efficient processing of large mbox files

## Supported Data Fields

- Email headers (Subject, From, To, Cc, Bcc, Reply-To)
- Email body content (plain text)
- Timestamps and Message IDs
- Reply chains and references
- Content length metrics

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   streamlit run app.py
   ```

2. Open your web browser to the displayed URL (typically http://localhost:8501)

3. Upload your mbox file using the file uploader

4. Select your desired output formats (Excel, CSV, TXT)

5. Click "Start Conversion" and download your converted files

## File Formats

### Input
- `.mbox` files (standard Unix mailbox format)

### Output
- **Excel (.xlsx)**: Structured spreadsheet with auto-sized columns
- **CSV (.csv)**: Comma-separated values for data analysis
- **TXT (.txt)**: Human-readable format with email separation

## Technical Details

- Built with Streamlit for the web interface
- Uses Python's built-in `mailbox` module for mbox parsing
- Pandas for data manipulation and Excel export
- openpyxl for Excel file generation
- Proper handling of MIME encoding and character sets

## Requirements

- Python 3.7+
- streamlit
- pandas
- openpyxl

## License

Open source - feel free to use and modify as needed.