@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

:: ============================================================================
:: Office Document Translator - Enhanced Edition v2.1
:: ============================================================================
:: Streamlined Office-Only Translation System
:: Supports: Excel, Word, PowerPoint  
:: Languages: Japanese, English, Vietnamese, Thai, Chinese, Korean
:: Powered by: Gemini 2.0 Flash API
:: ============================================================================

title Office Document Translator - Enhanced Edition

:: Clear screen and show header
cls
echo.
echo ===============================================================================
echo                    Office Document Translator                               
echo                      Enhanced Edition v2.1 - Streamlined                    
echo ===============================================================================
echo  Supports: Excel, Word, PowerPoint                                          
echo  Six Languages: Japanese, English, Vietnamese, Thai, Chinese, Korean        
echo  Optimized: No PDF Processing, Streamlined Dependencies, Fast Setup         
echo ===============================================================================
echo.

:: Check Python installation
echo [INFO] Verifying system requirements...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not accessible from PATH
    echo [HELP] Please install Python 3.8+ from https://python.org
    echo [HELP] Important: Check "Add Python to PATH" during installation
    echo [HELP] After installation, restart your terminal and try again
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo [OK] Python %PYTHON_VERSION% verified

:: Create directories
echo [INFO] Setting up workspace...
if not exist "input" (
    mkdir input >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo [OK] Created input directory
    ) else (
        echo [ERROR] Failed to create input directory
        exit /b 1
    )
) else (
    echo [OK] Input directory exists
)

if not exist "output" (
    mkdir output >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo [OK] Created output directory
    ) else (
        echo [ERROR] Failed to create output directory
        exit /b 1
    )
) else (
    echo [OK] Output directory exists
)

:: Install dependencies
echo [INFO] Preparing Enhanced Edition dependencies...
echo [SETUP] Installing optimized package set (no PDF dependencies)...

echo @echo off > temp_install.bat
echo python -m pip install --upgrade --quiet pip >> temp_install.bat
echo python -m pip install --upgrade --quiet openai python-dotenv >> temp_install.bat
echo python -m pip install --upgrade --quiet xlwings python-pptx python-docx >> temp_install.bat
echo python -m pip install --upgrade --quiet lxml Pillow pycryptodome zipfile36 pathlib rich tqdm >> temp_install.bat
echo if "%%OS%%"=="Windows_NT" python -m pip install --upgrade --quiet pywin32 comtypes >> temp_install.bat

call temp_install.bat >nul 2>&1
set INSTALL_RESULT=%ERRORLEVEL%
del temp_install.bat >nul 2>&1

if %INSTALL_RESULT% EQU 0 (
    echo [OK] Enhanced Edition dependencies installed successfully
) else (
    echo [WARN] Some dependencies may need manual attention
)

:: Verify components
echo [INFO] Verifying Enhanced Edition components...

python -c "import openai" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] Translation API client ready
) else (
    echo [WARN] Translation API client needs configuration
)

python -c "import xlwings" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] Excel processing engine ready
) else (
    echo [WARN] Excel processing in fallback mode
)

python -c "import docx" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] Word processing engine ready
) else (
    echo [WARN] Word processing in fallback mode
)

python -c "import pptx" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] PowerPoint processing engine ready
) else (
    echo [WARN] PowerPoint processing in fallback mode
)

echo.

:: Check API key
if not exist ".env" (
    echo [SETUP] First-time setup required
    echo.
    echo Please create a .env file with your Gemini API key:
    echo GEMINI_API_KEY=your_api_key_here
    echo.
    echo Get your FREE API key from: https://aistudio.google.com/app/apikey
    echo Enhanced Edition Guide: Supports 6 languages with optimized performance
    echo.
    echo After creating the .env file, run this script again
    pause
    exit /b 1
)

:: Ready to process files from input folder
echo [INFO] Ready to translate documents from input folder...

:: Language selection
echo.
echo ===============================================================================
echo                      Select Target Language - Enhanced Edition               
echo ===============================================================================
echo.
echo   1. Japanese (ja) - Business and Technical Documents
echo   2. Vietnamese (vi) - Southeast Asian Market
echo   3. English (en) - Global Business Standard
echo   4. Thai (th) - Thailand Market [NEW]
echo   5. Chinese Simplified (zh) - China Market [NEW]
echo   6. Korean (ko) - Korea Market [NEW]
echo.
echo Enhanced Edition supports all 6 languages with optimized performance
echo.
set /p "choice=Enter your choice (1-6): "

:: Process language choice
set "target_lang="
set "lang_description="
if "%choice%"=="1" (
    set "target_lang=ja"
    set "lang_description=Business and Technical Documents"
)
if "%choice%"=="2" (
    set "target_lang=vi"  
    set "lang_description=Southeast Asian Market"
)
if "%choice%"=="3" (
    set "target_lang=en"
    set "lang_description=Global Business Standard"
)
if "%choice%"=="4" (
    set "target_lang=th"
    set "lang_description=Thailand Market [NEW]"
)
if "%choice%"=="5" (
    set "target_lang=zh"
    set "lang_description=China Market [NEW]"
)
if "%choice%"=="6" (
    set "target_lang=ko"
    set "lang_description=Korea Market [NEW]"
)

if "%target_lang%"=="" (
    echo [ERROR] Invalid choice. Please select 1-6 and run the script again.
    echo [HELP] Enhanced Edition supports 6 languages for maximum market coverage
    pause
    exit /b 1
)

:: Language confirmation
set "lang_name="
if "%target_lang%"=="ja" set "lang_name=Japanese"
if "%target_lang%"=="vi" set "lang_name=Vietnamese"
if "%target_lang%"=="en" set "lang_name=English"
if "%target_lang%"=="th" set "lang_name=Thai"
if "%target_lang%"=="zh" set "lang_name=Chinese Simplified"
if "%target_lang%"=="ko" set "lang_name=Korean"

echo.
echo [OK] Target language: %lang_name%
echo [INFO] Focus: %lang_description%
echo.

:: Start translation
echo ===============================================================================
echo                        Starting Enhanced Translation                          
echo ===============================================================================
echo.
echo [INFO] Processing documents with Enhanced Edition engine...
echo [LANG] Target: %lang_name% (%lang_description%)
echo [TECH] Using Gemini 2.0 Flash API with optimized Office processing
echo.

:: Run the translator
python translator.py --dir input --to %target_lang%

:: Check results
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===============================================================================
    echo                    Enhanced Translation Complete!                          
    echo ===============================================================================
    echo.
    echo Translation Summary:
    echo   Target Language: %lang_name%
    echo   Documents Processed: Successfully completed translation
    echo   Output Location: output\ folder
    echo   Enhanced Edition: Optimized performance (no PDF overhead)
    echo.
    echo Enhanced Edition Benefits Delivered:
    echo   - Six-language support with Asian market focus
    echo   - Streamlined Office document processing  
    echo   - Faster performance with reduced dependencies
    echo.
    echo [INFO] Translated files are saved in the 'output' folder
)

echo.
echo ===============================================================================
echo              Thank you for using Enhanced Edition v2.1!                    
echo ===============================================================================
echo Press any key to exit...
pause >nul