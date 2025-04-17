@echo off
echo Cable Sizing Calculator
echo ====================

:: Install packages from requirements.txt
echo Installing required packages...
pip install -r requirements.txt

echo.
echo Starting Python FastAPI server...
echo.
echo Application will be available at http://localhost:8000
echo Press Ctrl+C to stop the server when done.
echo.

:: Start the browser in a separate process
start http://localhost:8000

:: Run the Python server in the foreground
:: The "call" command ensures the batch process waits for Python to complete
call python app.py

:: In case the Python server exits unexpectedly, keep the window open
pause

rem If you want to run the JavaScript version instead, open abcd.html directly in your browser.