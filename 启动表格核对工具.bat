@echo off
cd /d %~dp0
python excel_check_tool.py
if errorlevel 1 pause
