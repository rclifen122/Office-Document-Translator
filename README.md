# Office Document Translator

A powerful tool for translating Microsoft Office documents (Excel, Word, PowerPoint) between Japanese, English, and Vietnamese while preserving formatting and structure.

## Features

- **Multi-format support**: Translate Excel (.xlsx, .xls), Word (.docx, .doc), and PowerPoint (.pptx, .ppt) files
- **Tri-lingual support**: Translate between Japanese, English, and Vietnamese
- **Format preservation**: Maintains original document formatting, styles, and layout
- **Rich progress UI**: Real-time progress tracking with detailed status updates
- **Batch processing**: Process multiple files in one operation
- **Smart text detection**: Automatically identifies text that needs translation
- **Robust error handling**: Recovers from API failures with automatic retries
- **Dependency management**: Automatically installs required packages

## Requirements

- Python 3.7 or higher
- Windows OS (for batch file execution)
- Gemini API key (obtain from [Google AI Studio](https://aistudio.google.com/app/apikey))

## Installation

1. **Clone or download this repository** to your local machine

2. **Set up your API key:**
   - Create a `.env` file in the project directory (or run the batch file once to create it)
   - Add your Gemini API key: `GEMINI_API_KEY=your_api_key_here`

3. **Install dependencies:**
   Dependencies will be installed automatically when you run the batch file, or you can install them manually:
   ```
   pip install -r translator-requirements.txt
   ```

## Usage

### Option 1: Using the Batch File (Recommended for Windows)

1. Place your Office documents (Excel, Word, PowerPoint) in the `input` folder
2. Run `run_translator.bat`
3. Choose the translation direction:
   - Option 1: Vietnamese to Japanese
   - Option 2: Japanese to Vietnamese
   - Option 3: To English
4. Wait for the translation to complete
5. Find your translated documents in the `output` folder

### Option 2: Using Command Line

```
python translator.py [OPTIONS]
```

Options:
- `--to`: Target language (`ja` for Japanese, `vi` for Vietnamese, `en` for English). Default: `ja`
- `--file`: Path to a specific file to translate (optional)
- `--dir`: Path to a directory containing files to translate (optional)
- `--output-dir`: Path to output directory (optional, default: `./output`)
- `--version`: Show version information

Examples:
```
# Translate all supported files in the input directory to Japanese
python translator.py

# Translate a specific file to Vietnamese
python translator.py --to vi --file path/to/your/document.xlsx

# Translate a specific file to English
python translator.py --to en --file path/to/your/document.xlsx

# Translate all files in a specific directory to Japanese
python translator.py --to ja --dir path/to/your/folder

# Specify a custom output directory
python translator.py --output-dir path/to/custom/output
```

## Supported Content Types

The translator processes the following content within documents:

### Excel Files
- Cell content
- Text in shapes
- WordArt text
- Text in embedded objects

### PowerPoint Files
- Text in slides (titles, content)
- Text in shapes
- Table content
- Text in groups
- Speaker notes

### Word Files
- Paragraph text
- Table content
- Headers and footers
- Content in all document sections

## Troubleshooting

### Missing API Key
If you see a warning about a missing API key:
1. Open the `.env` file in the project directory
2. Replace `your_api_key_here` with your actual Gemini API key
3. Save the file and run the translator again

### Installation Errors
If you encounter errors during package installation:
1. Ensure you have Python 3.7+ installed and in your PATH
2. Try installing the packages manually: `pip install -r translator-requirements.txt`
3. Check if you have administrator permissions if required

### File Processing Errors
If a file fails to translate:
1. Ensure the file is not open in another application
2. Verify the file is not password-protected
3. Confirm the file format is supported (.xlsx, .xls, .docx, .doc, .pptx, .ppt)

## How It Works

1. The translator scans each document for text content
2. Text segments are extracted and sent to the Gemini AI API for translation
3. The translated text is inserted back into the original document
4. The document is saved with formatting intact in the output directory

## License

This project is provided as-is for educational and practical purposes.

## Acknowledgments

This project is based on [ai-excel-translator](https://github.com/hoangduong92/ai-excel-translator) by [hoangduong92](https://github.com/hoangduong92). The original Excel translation functionality has been extended to support additional file formats (Word, PowerPoint) while maintaining the core translation approach.

Special thanks to:
- **hoangduong92** for creating the original AI-powered Excel translator that inspired this project
- The original project supports various language pairs and works with different AI providers including Gemini, OpenAI, and Azure

Please respect the license terms of the original project when using, modifying, or distributing this software.
