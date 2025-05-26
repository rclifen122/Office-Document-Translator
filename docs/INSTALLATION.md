# 📦 Installation Guide

This guide will walk you through installing Office Document Translator on your system.

## 🔧 System Requirements

### Minimum Requirements
- **Operating System**: Windows 7/8/10/11 (64-bit recommended)
- **Python**: 3.7 or higher
- **Memory**: 4GB RAM minimum, 8GB recommended
- **Storage**: 500MB free space for installation
- **Internet**: Required for API access and package installation

### Recommended Requirements
- **Operating System**: Windows 10/11 (64-bit)
- **Python**: 3.9 or 3.10
- **Memory**: 8GB RAM or more
- **Storage**: 2GB free space
- **Internet**: Stable broadband connection

## 🚀 Quick Installation

### Option 1: Using Git (Recommended)

```bash
# Clone the repository
git clone https://github.com/rclifen122/Office-Document-Translator.git

# Navigate to the project directory
cd Office-Document-Translator

# Install dependencies
pip install -r translator-requirements.txt

# Set up environment
copy .env.example .env
# Edit .env with your API key
```

### Option 2: Download ZIP

1. Go to [GitHub repository](https://github.com/rclifen122/Office-Document-Translator)
2. Click "Code" → "Download ZIP"
3. Extract the ZIP file to your desired location
4. Open Command Prompt/PowerShell in the extracted folder
5. Run: `pip install -r translator-requirements.txt`

## 🔑 API Key Setup

### Getting Your Gemini API Key

1. Visit [Google AI Studio](https://aistudio.google.com/app/apikey)
2. Sign in with your Google account
3. Click "Create API Key"
4. Copy the generated key

### Configuring the API Key

1. **Option A: Using .env file (Recommended)**
   ```bash
   # Copy the example file
   copy .env.example .env
   
   # Edit .env file and add your key
   GEMINI_API_KEY=your_actual_api_key_here
   ```

2. **Option B: Environment Variable**
   ```bash
   # Windows Command Prompt
   set GEMINI_API_KEY=your_actual_api_key_here
   
   # Windows PowerShell
   $env:GEMINI_API_KEY="your_actual_api_key_here"
   ```

## 🐍 Python Installation

### If Python is Not Installed

1. **Download Python**: Visit [python.org](https://www.python.org/downloads/)
2. **Install Python**: 
   - ✅ Check "Add Python to PATH"
   - ✅ Check "Install for all users" (if you have admin rights)
3. **Verify Installation**:
   ```bash
   python --version
   pip --version
   ```

### Managing Multiple Python Versions

If you have multiple Python versions:

```bash
# Use specific Python version
python3.9 -m pip install -r translator-requirements.txt

# Or create a virtual environment
python -m venv translator-env
translator-env\Scripts\activate  # Windows
pip install -r translator-requirements.txt
```

## 📦 Dependency Installation

### Automatic Installation

Run the batch file and dependencies will be installed automatically:
```bash
scripts\run_translator.bat
```

### Manual Installation

```bash
# Basic dependencies
pip install -r translator-requirements.txt

# For building executables
pip install -r requirements_exe.txt

# Development dependencies (optional)
pip install flake8 pytest pytest-cov
```

### Common Installation Issues

**Issue: Permission Denied**
```bash
# Solution: Use --user flag
pip install --user -r translator-requirements.txt
```

**Issue: SSL Certificate Error**
```bash
# Solution: Upgrade pip and certificates
python -m pip install --upgrade pip
pip install --upgrade certifi
```

**Issue: Package Conflicts**
```bash
# Solution: Use virtual environment
python -m venv translator-env
translator-env\Scripts\activate
pip install -r translator-requirements.txt
```

## ✅ Verification

### Test Your Installation

1. **Check Python and packages**:
   ```bash
   python --version
   python -c "import openpyxl, python_docx, python_pptx; print('All packages installed!')"
   ```

2. **Test the translator**:
   ```bash
   python translator.py --version
   ```

3. **Test the GUI**:
   ```bash
   python gui_translator.py
   ```

### Sample Test

1. Create a simple Excel file with some text
2. Place it in the `input/` folder
3. Run: `scripts\run_translator.bat`
4. Check the `output/` folder for translated file

## 🔧 Troubleshooting

### Common Problems

**Problem**: "Python is not recognized"
- **Solution**: Add Python to your PATH or reinstall with "Add to PATH" checked

**Problem**: "No module named 'openpyxl'"
- **Solution**: Run `pip install -r translator-requirements.txt`

**Problem**: API key not working
- **Solution**: Verify your API key is correct and active

**Problem**: Translation fails
- **Solution**: Check internet connection and API quotas

### Getting Help

- 📧 **Issues**: [GitHub Issues](https://github.com/rclifen122/Office-Document-Translator/issues)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/rclifen122/Office-Document-Translator/discussions)
- 📖 **Documentation**: [Project Wiki](https://github.com/rclifen122/Office-Document-Translator/wiki)

## 🚀 Next Steps

After installation:
1. 📚 Read the [Usage Guide](USAGE.md)
2. 🎯 Try the [Quick Start Examples](EXAMPLES.md)
3. ⚙️ Explore [Advanced Configuration](CONFIGURATION.md)

---

✅ **Installation Complete!** You're ready to start translating documents! 