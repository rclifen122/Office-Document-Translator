#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Build Script for Office Document Translator Executable
=====================================================
Automates the creation of a standalone executable using PyInstaller.
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def install_pyinstaller():
    """Install PyInstaller if not available"""
    print("🔧 Installing PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller>=5.0.0"])
        print("✅ PyInstaller installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install PyInstaller: {e}")
        return False

def create_spec_file():
    """Create custom PyInstaller spec file"""
    print("📄 Creating PyInstaller spec file...")
    
    spec_content = """
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Hidden imports for Office document processing
hidden_imports = [
    'xlwings',
    'xlwings.conversion',
    'xlwings.constants',
    'pptx',
    'pptx.presentation',
    'pptx.slide',
    'docx',
    'docx.document',
    'openai',
    'dotenv',
    'rich',
    'rich.console',
    'rich.progress',
    'tqdm',
    'lxml',
    'lxml.etree',
    'PIL',
    'PIL.Image',
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
    'pathlib',
    'threading',
    'webbrowser',
    'datetime',
    'zipfile',
    'xml.etree.ElementTree',
    'comtypes',
    'comtypes.client',
    'win32com',
    'win32com.client',
    'pycryptodome',
    'Crypto',
    'Crypto.Cipher'
]

# Data files to include
datas = [
    ('translator.py', '.'),
    ('translator-system-prompt.txt', '.'),
]

a = Analysis(
    ['gui_translator.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OfficeTranslator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon here if available
    version='file_version_info.txt'  # Add version info if available
)
"""

    with open("OfficeTranslator.spec", "w", encoding="utf-8") as f:
        f.write(spec_content.strip())
    
    print("✅ Spec file created: OfficeTranslator.spec")

def create_version_info():
    """Create version information file"""
    print("📋 Creating version info...")
    
    version_info = """
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(2, 1, 0, 0),
    prodvers=(2, 1, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          '040904B0',
          [
            StringStruct('CompanyName', 'Office Document Translator'),
            StringStruct('FileDescription', 'Enhanced Edition Office Document Translator'),
            StringStruct('FileVersion', '2.1.0.0'),
            StringStruct('InternalName', 'OfficeTranslator'),
            StringStruct('LegalCopyright', 'Enhanced Edition v2.1'),
            StringStruct('OriginalFilename', 'OfficeTranslator.exe'),
            StringStruct('ProductName', 'Office Document Translator - Enhanced Edition'),
            StringStruct('ProductVersion', '2.1.0.0')
          ]
        )
      ]
    ),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""

    with open("file_version_info.txt", "w", encoding="utf-8") as f:
        f.write(version_info.strip())
    
    print("✅ Version info created: file_version_info.txt")

def build_executable():
    """Build the executable using PyInstaller"""
    print("🚀 Building executable...")
    
    try:
        # Clean previous builds
        if os.path.exists("build"):
            shutil.rmtree("build")
        if os.path.exists("dist"):
            shutil.rmtree("dist")
        
        # Build using spec file
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller", 
            "--clean", 
            "OfficeTranslator.spec"
        ])
        
        print("✅ Executable built successfully!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        return False

def create_distribution_package():
    """Create a complete distribution package"""
    print("📦 Creating distribution package...")
    
    # Create distribution directory
    dist_dir = Path("OfficeTranslator_v2.1")
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    
    dist_dir.mkdir()
    
    try:
        # Copy executable
        exe_source = Path("dist/OfficeTranslator.exe")
        if exe_source.exists():
            shutil.copy2(exe_source, dist_dir / "OfficeTranslator.exe")
        else:
            print("❌ Executable not found in dist folder")
            return False
        
        # Create directories
        (dist_dir / "input").mkdir()
        (dist_dir / "output").mkdir()
        (dist_dir / "setup").mkdir()
        
        # Create README
        readme_content = """# Office Document Translator - Enhanced Edition v2.1

## Quick Start

1. **Double-click** `OfficeTranslator.exe` to launch the application
2. **Setup API Key** (first time only):
   - Click "Setup API Key" button
   - Get your FREE Gemini API key from the link provided
   - Paste your API key and click Save
3. **Select Files**:
   - Place your Office documents in the `input` folder, OR
   - Use "Browse..." to select a different input folder
4. **Choose Language** from the dropdown menu
5. **Click "Start Translation"** and wait for completion
6. **Find Results** in the `output` folder

## Supported Files
- Excel: .xlsx, .xls
- Word: .docx, .doc  
- PowerPoint: .pptx, .ppt

## Supported Languages
- 🇯🇵 Japanese
- 🇻🇳 Vietnamese
- 🇬🇧 English
- 🇹🇭 Thai
- 🇨🇳 Chinese (Simplified)
- 🇰🇷 Korean

## Features
- Preserves original formatting
- Handles complex elements (charts, tables, shapes)
- Batch processing support
- Professional translation quality
- User-friendly interface

## Requirements
- Windows 10/11 (64-bit)
- Internet connection for translation
- FREE Gemini API key

## Support
For issues or questions, check the Activity Log in the application.
"""

        with open(dist_dir / "README.txt", "w", encoding="utf-8") as f:
            f.write(readme_content)
        
        # Create API setup guide
        api_setup_content = """# Getting Your FREE Gemini API Key

## Step 1: Visit Google AI Studio
Open this link in your web browser:
https://aistudio.google.com/app/apikey

## Step 2: Sign in with Google Account
Use your existing Google account or create a new one (free).

## Step 3: Create API Key
1. Click "Create API Key"
2. Select "Create API key in new project" (recommended)
3. Copy the generated API key

## Step 4: Add to Office Translator
1. Open Office Translator application
2. Click "Setup API Key" button
3. Paste your API key
4. Click "Save"

## Important Notes
- The API key is FREE with generous usage limits
- Your key is stored locally and securely
- Never share your API key with others
- You can regenerate the key if needed

## Troubleshooting
- If the link doesn't work, search for "Google AI Studio" in your browser
- Make sure you're signed in to your Google account
- The API key should be a long string of letters and numbers
"""

        with open(dist_dir / "setup" / "API_KEY_SETUP.txt", "w", encoding="utf-8") as f:
            f.write(api_setup_content)
        
        # Create supported formats guide
        formats_content = """# Supported File Formats

## Excel Files
- .xlsx (Excel 2007 and later)
- .xls (Excel 97-2003)

Translates:
- Cell content
- Chart titles and labels
- Shape text
- Comments and notes

## Word Documents  
- .docx (Word 2007 and later)
- .doc (Word 97-2003)

Translates:
- Document text
- Headers and footers
- Table content
- Text boxes

## PowerPoint Presentations
- .pptx (PowerPoint 2007 and later)
- .ppt (PowerPoint 97-2003)

Translates:
- Slide text
- Title and content placeholders
- Table content
- Notes section
- Basic shapes with text

## What Gets Preserved
- Original formatting (fonts, colors, styles)
- Document structure and layout
- Images and graphics (untranslated)
- Formulas and calculations (Excel)
- Hyperlinks and references

## File Size Limits
- Individual files: Up to 100MB
- Batch processing: No limit on number of files
- Large files may take longer to process

## Tips for Best Results
- Use clear, well-structured documents
- Avoid heavily stylized or artistic text
- Check translated output for accuracy
- Keep backups of original files
"""

        with open(dist_dir / "setup" / "SUPPORTED_FORMATS.txt", "w", encoding="utf-8") as f:
            f.write(formats_content)
        
        # Create sample .env template
        env_template = """# Office Document Translator Configuration
# Replace 'your_api_key_here' with your actual Gemini API key

GEMINI_API_KEY=your_api_key_here
"""

        with open(dist_dir / "setup" / "env_template.txt", "w", encoding="utf-8") as f:
            f.write(env_template)
        
        print(f"✅ Distribution package created: {dist_dir}")
        print(f"📁 Package contents:")
        print(f"   - OfficeTranslator.exe (main application)")
        print(f"   - input/ (place files here)")
        print(f"   - output/ (translated files)")
        print(f"   - README.txt (quick start guide)")
        print(f"   - setup/ (detailed setup guides)")
        
        return True
        
    except Exception as e:
        print(f"❌ Failed to create distribution package: {e}")
        return False

def main():
    """Main build process"""
    print("🏗️ Office Document Translator - Executable Builder")
    print("=" * 50)
    
    # Check if we're in the right directory
    required_files = ["gui_translator.py", "translator.py"]
    missing_files = [f for f in required_files if not os.path.exists(f)]
    
    if missing_files:
        print(f"❌ Missing required files: {missing_files}")
        print("Please run this script from the project root directory.")
        return 1
    
    # Step 1: Install PyInstaller
    if not install_pyinstaller():
        return 1
    
    # Step 2: Create configuration files
    create_version_info()
    create_spec_file()
    
    # Step 3: Build executable
    if not build_executable():
        return 1
    
    # Step 4: Create distribution package
    if not create_distribution_package():
        return 1
    
    print("\n🎉 BUILD COMPLETE!")
    print("=" * 50)
    print("✅ Your executable is ready in: OfficeTranslator_v2.1/")
    print("📋 Share the entire folder with users")
    print("🚀 Users can run OfficeTranslator.exe directly")
    print("\n💡 Next steps:")
    print("   1. Test the executable on a clean Windows machine")
    print("   2. Zip the OfficeTranslator_v2.1 folder for easy distribution")
    print("   3. Provide users with the README.txt for setup instructions")
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 