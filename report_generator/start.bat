@echo off
REM 检查Python的安装路径
set PYTHON=python
set PYTHON3=python3

REM 优先使用Python3
if exist %PYTHON3% (
    set PYTHON=%PYTHON3%
)

REM 使用Python运行demo.py
%PYTHON% ReportGenX.py

REM 检查是否成功运行
if %ERRORLEVEL% neq 0 (
    echo Failed to run ReportGenX.py
    pause
    exit /b 1
)

echo Successfully ran ReportGenX.py
pause
exit /b 0
