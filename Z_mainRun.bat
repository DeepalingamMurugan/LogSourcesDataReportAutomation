@echo off

REM Activate the virtual environment
call "Drive:\..\..\Scripts\activate.bat"

REM Path to the index page
set Index_path="Drive\..\..\indexStreamApp.py"

REM Py cmd to run the index streamlit
python -m streamlit run %Index_path%

REM Pause to keep the command prompt window open
pause