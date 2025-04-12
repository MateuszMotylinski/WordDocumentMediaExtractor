@echo off
REM Set script and icon file names
set SCRIPT=../Source/WordExtractorRunner.py
set ICON=WordExtractorIcon.ico
set EXENAME=WordExtractor

REM Clean up old build/dist folders
rmdir /s /q build
rmdir /s /q dist


REM Run PyInstaller using Python's -m option
python -m PyInstaller --onefile --icon=%ICON% --name=%EXENAME% %SCRIPT%

pause