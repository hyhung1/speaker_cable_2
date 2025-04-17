# Cable Sizing Calculator

A professional web application to calculate voltage drop, voltage at speaker, and SPL reduction for cable installations in audio systems.

## Features

- Clean, modern interface with company logo
- Input the number of speakers
- Specify default cable type (cross-sectional area in mm²) for all cables
- Input individual cable length (m) and power tapping (W) for each cable
- Calculate voltage drop at each speaker using the formula: 1.68 * 10^-8 * cable length * Power Tapping / (100 * cable type * 10^-6)
- Calculate voltage at speaker (Vs = previous speaker voltage - voltage drop)
- Calculate SPL reduction at each speaker (SPL = 20log(Vs/100))
- Display all results with 3 decimal places for precision
- Export calculations as Excel with a single click
- **Upload existing Excel files** to continue working on previous calculations
- Status notification system for user feedback
- Color-coded results (green for normal, red for abnormal)

## Setup Options

### JavaScript Version (Original)

Simply open `abcd.html` directly in your web browser - no server or installation required!

### Python Version (FastAPI)

1. Install Python 3.9 or later
2. Install the required dependencies: `pip install -r requirements.txt`
3. Run the application: `python app.py`
4. Access the application at [http://localhost:8000](http://localhost:8000)

Alternatively, you can use the batch file for a one-click setup:
1. Run `run_calculator.bat`
2. The batch file will:
   - Install all required dependencies
   - Start the FastAPI server
   - Open your default browser to the application URL

## How to Use

1. Open the application in your web browser
2. The application loads with 12 predefined cables with data from the provided example
3. You can change the number of cables and default cable type if needed
4. Click "Generate Table" to create a new table with the specified number of rows
5. Input or modify the cable type, length, and power tapping for each cable
6. Click "Calculate" to see the results
7. Click "Download as Excel" to export the calculation to an Excel file named 'cable_sizing.xlsx'
8. Use "Reset" to clear all results and start over

### Using the Excel Upload Feature

1. Click the "Upload Excel File" button in the upload section
2. Select a previously saved Excel file from your calculator
3. Click "Load Data" to import the data into the application
4. The table will automatically populate with the loaded data
5. You can continue editing the values or recalculate as needed

## Technical Details

### Calculation Method

The application uses the following formulas:
- Voltage Drop (Vd) = 1.68 * 10^-8 * cable length * Power Tapping / (100 * cable type * 10^-6)
- Voltage at Speaker (Vs):
  - For first cable: 100V - Voltage Drop
  - For subsequent cables: Previous Speaker Voltage - Voltage Drop
- SPL Reduction = 20 * log10(Voltage at Speaker / 100)

### Examine Status

The application automatically determines if the SPL reduction is within acceptable limits:
- **Normal** (Green): SPL reduction > -2 dB (acceptable power loss)
- **Abnormal** (Red): SPL reduction ≤ -2 dB (unacceptable power loss)

This is especially important for alarm and evacuation scenarios where maximum power loss should not exceed -2 dB.

### Excel Export Process

The Excel generation process works as follows:
1. When you click "Download as Excel", the application collects all table data
2. If the template file (cable2.xlsx) exists, it's used as a base to preserve formatting
3. Data is written to the appropriate cells with proper formatting
4. Excel file is generated and downloaded as 'cable_sizing.xlsx'
5. The file includes all data with proper formatting and color coding for status

## Project Structure

### JavaScript Version
- `abcd.html` - Main HTML file with the user interface and JavaScript functionality
- `css/styles.css` - Styling for the application
- `vector_logo1.png` - Company logo displayed in the header

### Python Version (FastAPI)
- `app.py` - Python backend with FastAPI server and all route handlers
- `templates/index.html` - HTML template file for the web interface
- `static/vector_logo1.png` - Company logo
- `cable2.xlsx` - Excel template file used for formatting exported files
- `requirements.txt` - Python dependencies (FastAPI, Pandas, Openpyxl, etc.)
- `run_calculator.bat` - Batch file for easy setup and execution on Windows

## Dependencies (Python Version)

- FastAPI - Web framework for building APIs
- Pandas - Data manipulation and analysis
- Openpyxl - Excel file handling
- Uvicorn - ASGI server implementation
- Jinja2 - Template engine
- Python-multipart - Multipart form parser
- Aiofiles - Asynchronous file operations
- Pydantic - Data validation and settings management

## Version Comparison

| Feature | JavaScript Version | Python Version |
|---------|-------------------|---------------|
| Server Required | No (browser only) | Yes (Python/FastAPI) |
| Installation | None needed | Python + dependencies |
| Calculations | Client-side | Server-side |
| Excel Generation | Client-side (SheetJS) | Server-side (Openpyxl) |
| Excel Template Support | Limited | Full (preserves formatting) |
| Distribution | Just copy files | Requires Python setup | 