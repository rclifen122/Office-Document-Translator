@echo off
:: Office Document Translator - GUI Launcher
:: For users who want to try the modern interface

title Office Document Translator - GUI Version

echo.
echo ===============================================================================
echo                    Office Document Translator - GUI Version                    
echo ===============================================================================
echo.
echo  🖱️ Starting modern graphical interface...
echo  📁 Input/Output folders will be created automatically
echo  🔑 API key setup available in the interface
echo.

python gui_translator.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ GUI launch failed. Possible issues:
    echo   - Python not installed or not in PATH
    echo   - Missing dependencies (run: pip install -r requirements_exe.txt)
    echo   - GUI libraries not available
    echo.
    echo 💡 Try the batch version instead: run_translator.bat
    echo.
    pause
) 