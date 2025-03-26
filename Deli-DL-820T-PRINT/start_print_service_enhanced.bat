@echo off

e:
cd e:\github\InterOfHardware\Deli-DL-820T-PRINT

:: 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    exit /b 1
)

:: 检查必要的模块
python -c "import win32print, win32ui, qrcode, flask" >nul 2>&1
if %errorlevel% neq 0 (
    pip install pywin32 qrcode flask pillow >nul 2>&1
)

:: 启动服务
pythonw direct_print.py