# 📄 Office Document Translator

[![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![AI](https://img.shields.io/badge/AI-Google%20Gemini-orange.svg)](https://ai.google.dev/)

> A powerful, AI-driven tool for translating Microsoft Office documents (Excel, Word, PowerPoint) between Japanese, English, and Vietnamese while perfectly preserving formatting and structure.

## 🌟 Features

### 🔧 Core Capabilities
- **📊 Multi-format Support**: Excel (.xlsx, .xls), Word (.docx, .doc), PowerPoint (.pptx, .ppt)
- **🌐 Tri-lingual Translation**: Japanese ↔ English ↔ Vietnamese
- **🎨 Format Preservation**: Maintains original formatting, styles, layouts, and structures
- **⚡ Batch Processing**: Process multiple files simultaneously
- **🤖 Smart Text Detection**: Automatically identifies translatable content

### 🖥️ User Experience
- **🎯 GUI Interface**: User-friendly graphical interface (`gui_translator.py`)
- **📊 Real-time Progress**: Live progress tracking with detailed status updates
- **🔄 Robust Error Handling**: Automatic retries and recovery from API failures
- **📦 Auto Dependencies**: Automatic package installation and management

### 🛠️ Developer Features
- **🔨 Build Tools**: Create standalone executables with `build_exe.py`
- **⚙️ Command Line Interface**: Full CLI support for automation
- **📝 Comprehensive Logging**: Detailed operation logs for debugging

## 🚀 Quick Start

### Prerequisites
- **Python 3.7+** ([Download](https://www.python.org/downloads/))
- **Windows OS** (for batch file execution)
- **Google Gemini API Key** ([Get yours here](https://aistudio.google.com/app/apikey))

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/rclifen122/Office-Document-Translator.git
   cd Office-Document-Translator
   ```

2. **Set up API credentials**
   ```bash
   # Create .env file
   echo GEMINI_API_KEY=your_api_key_here > .env
   ```

3. **Install dependencies**
   ```bash
   pip install -r translator-requirements.txt
   ```

## 📖 Usage

### 🖱️ GUI Mode (Recommended)
```bash
python gui_translator.py
# or
scripts\launch_gui.bat
```

### 🖥️ Command Line Mode
```bash
# Basic usage
python translator.py

# Translate specific file
python translator.py --file document.xlsx --to ja

# Batch translate directory
python translator.py --dir ./documents --to en --output-dir ./translated
```

### 🔄 Batch Mode (Windows)
1. 📁 Place documents in the `input/` folder
2. ▶️ Run `scripts\run_translator.bat`
3. 🎯 Select translation direction
4. ✅ Get results in `output/` folder

## 🎯 Supported File Types & Content

| File Type | Extensions | Supported Content |
|-----------|------------|------------------|
| **Excel** | `.xlsx`, `.xls` | Cell content, shapes, WordArt, embedded objects |
| **Word** | `.docx`, `.doc` | Paragraphs, tables, headers/footers, all sections |
| **PowerPoint** | `.pptx`, `.ppt` | Slide content, shapes, tables, speaker notes |

## ⚙️ Configuration Options

### Command Line Arguments
```bash
--to LANG          # Target language: ja, en, vi
--file PATH        # Single file to translate
--dir PATH         # Directory to process
--output-dir PATH  # Output directory
--version          # Show version info
```

### Environment Variables
```bash
GEMINI_API_KEY=your_key_here  # Required: Your Google Gemini API key
```

## 🔧 Building Executables

Create standalone executables for distribution:

```bash
python build_exe.py
```

This generates:
- 📦 `dist/OfficeTranslator.exe` - Standalone executable
- 📁 Complete package with all dependencies

## 🐛 Troubleshooting

### 🔑 API Key Issues
```bash
# Verify your .env file
cat .env
# Should show: GEMINI_API_KEY=your_actual_key
```

### 📦 Installation Problems
```bash
# Update pip and try again
pip install --upgrade pip
pip install -r translator-requirements.txt
```

### 📄 File Processing Errors
- ✅ Ensure files are not open in other applications
- ✅ Check file is not password-protected
- ✅ Verify supported file format
- ✅ Check file permissions

## 🏗️ Project Structure

```
Office-Document-Translator/
├── 📄 translator.py              # Core translation engine
├── 🖥️ gui_translator.py          # GUI interface
├── 🔨 build_exe.py               # Executable builder
├── 📋 translator-requirements.txt # Dependencies
├── 📦 requirements_exe.txt       # Build dependencies
├── 📁 scripts/                   # Batch files and utilities
│   ├── ⚙️ run_translator.bat     # Windows batch runner
│   └── 🚀 launch_gui.bat         # GUI launcher
├── 📁 docs/                      # Documentation
│   └── 📖 INSTALLATION.md        # Installation guide
├── 📁 .github/                   # GitHub configurations
│   ├── 🔧 workflows/             # CI/CD workflows
│   └── 📝 ISSUE_TEMPLATE/        # Issue templates
├── 📁 input/                     # Input documents folder
├── 📁 output/                    # Translated documents folder
├── 📄 README.md                  # This file
├── 🤝 CONTRIBUTING.md            # Contributing guidelines
├── 🔒 SECURITY.md                # Security policy
└── 📋 CHANGELOG.md               # Version history
```

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details.

1. 🍴 Fork the repository
2. 🌿 Create a feature branch
3. 💻 Make your changes
4. ✅ Add tests if applicable
5. 📝 Update documentation
6. 🚀 Submit a pull request

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

This project builds upon the excellent work of:
- **[hoangduong92](https://github.com/hoangduong92)** - Original [ai-excel-translator](https://github.com/hoangduong92/ai-excel-translator)
- **Google AI** - Gemini API for translation services
- **Microsoft** - Office document format specifications

## 📊 Version History

| Version | Date | Changes |
|---------|------|---------|
| 2.0.0 | 2024-05 | Added GUI, multi-format support, build tools |
| 1.0.0 | 2024-04 | Initial release with Excel support |

## 🆘 Support

- 📧 **Issues**: [GitHub Issues](https://github.com/rclifen122/Office-Document-Translator/issues)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/rclifen122/Office-Document-Translator/discussions)
- 📖 **Documentation**: [Project Wiki](https://github.com/rclifen122/Office-Document-Translator/wiki)

---

<div align="center">

**⭐ Star this repository if you find it useful! ⭐**

Made with ❤️ for the global translation community

</div>
