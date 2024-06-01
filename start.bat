@echo off
rem 隐藏黑色窗口
if "%1"=="h" goto begin 
mshta vbscript:createobject("wscript.shell").run("%~nx0 h",0)(window.close)&&exit 
:begin 
REM 检查Python的安装路径
set PYTHON=python
set PYTHON3=python3

REM 优先使用Python3
if exist %PYTHON3% (
    set PYTHON=%PYTHON3%
)

REM 使用Python运行demo.py
%PYTHON% ShitReport.py

REM 检查是否成功运行
if %ERRORLEVEL% neq 0 (
    echo Failed to run ShitReport.py
    pause
    exit /b 1
)

echo Successfully ran ShitReport.py
pause
exit /b 0
