@echo off

echo Uploaded file: %UPLOAD_FILE% 

REM Rename uploaded file to its original name 
rename UPLOAD_FILE "%UPLOAD_FILE%"

REM Verify file is present 
dir

REM Delete all files inside Result folder BEFORE running Python
if exist Result (
    echo Cleaning Result folder...
    del /q Result\*
) else (
    mkdir Result
)


REM Run Python conversion script
python excel_to_xml.py "%UPLOAD_FILE%"

REM Delete Excel files older than 7 days
echo Deleting Excel files older than 7 days...
forfiles /m *.xlsx /d -7 /c "cmd /c del @file"

REM Verify remaining files
dir
