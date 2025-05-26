# System Patterns: Office Document Translator - Enhanced Edition

## Architecture Overview

### Enhanced High-Level Design
```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   Dual UI       │───▶│  File Processor  │───▶│  Output Manager │
│ (CLI + GUI)     │    │   (Office Only)  │    │                 │
└─────────────────┘    └──────────────────┘    └─────────────────┘
         │                       │
         │                       ▼
         │           ┌──────────────────────┐
         │           │  Enhanced Translation│
         └──────────▶│  Engine (6 Languages)│
                     │  + System Prompt     │
                     └──────────────────────┘
                                │
                    ┌───────────┼───────────┐
                    ▼           ▼           ▼
            ┌──────────┐ ┌──────────┐ ┌──────────┐
            │   Excel  │ │   Word   │ │PowerPoint│
            │Processor │ │Processor │ │ Advanced │
            │          │ │          │ │ Processor │
            └──────────┘ └──────────┘ └──────────┘
```

### Enhanced Core Components

#### 1. Dual Interface Architecture - NEW
- **Pattern**: Multi-interface facade with shared backend
- **CLI Interface**: Enhanced colorful batch file for technical users
- **GUI Interface**: Modern tkinter application for non-technical users
- **Shared Engine**: Both interfaces use the same translation engine
- **Executable Distribution**: PyInstaller build for zero-dependency deployment

#### 2. System Prompt Architecture - NEW
- **Pattern**: External configuration pattern for AI model instructions
- **File Location**: `translator-system-prompt.txt` in project root
- **Loading Strategy**: Dynamic runtime loading with fallback creation
- **API Integration**: Used as system message in Gemini API calls
- **Customization**: Easy modification without code changes

#### 3. Enhanced File Type Detection System
- **Pattern**: Strategy Pattern for Office document routing
- **Implementation**: DocType enum focused on Office formats
- **Location**: `get_file_type()` function
- **Supported Types**: Excel (.xlsx, .xls), Word (.docx, .doc), PowerPoint (.pptx, .ppt)
- **Purpose**: Route Office files to appropriate specialized processors

#### 4. Advanced PowerPoint Processing Architecture (Maintained)
- **Pattern**: Chain of Responsibility with fallback mechanisms
- **Implementation**: AdvancedPowerPointProcessor with multiple engines
- **Engines**: Hybrid → Enhanced PPTX → COM Automation → XML Direct
- **Purpose**: Maximum element coverage with graceful degradation
- **Focus**: Complex Office elements (SmartArt, WordArt, OLE objects)

#### 5. Enhanced Translation Pipeline
- **Pattern**: Pipeline Pattern with six-language batch processing
- **Languages**: Japanese, Vietnamese, English, Thai, Chinese (Simplified), Korean
- **Stages**: Extract → Batch → Translate → Apply
- **Optimization**: Configurable batch sizes (default: 50 items)
- **Error Handling**: Retry logic with exponential backoff
- **System Prompt Integration**: Dynamic instruction loading for customizable behavior

#### 6. Modern User Interface Architecture
- **Pattern**: Facade Pattern with hidden complexity
- **CLI Implementation**: Professional colorful batch file interface
- **GUI Implementation**: Modern tkinter with threading architecture
- **Features**: ANSI colors, Unicode support, visual progress indicators
- **Installation**: Silent pip operations with progress-only display
- **User Experience**: Interactive language selection with flag emojis

## Key Technical Decisions

### 1. Dual Interface Strategy (Major Addition)
**Decision**: Provide both CLI and GUI interfaces for different user types
**Rationale**: Serve both technical users (CLI) and non-technical users (GUI)
**Benefits**: Broader market appeal, professional deployment options, user choice
**Implementation**: Shared backend with separate interface modules

### 2. System Prompt Architecture (Major Discovery)
**Decision**: Use external file for AI model instructions
**Rationale**: Flexibility, maintainability, easy customization without code changes
**Benefits**: Easy language addition, A/B testing, version control of prompts
**Implementation**: Dynamic loading in translate_batch() with fallback system

### 3. GUI Framework Selection
**Decision**: Use tkinter for GUI implementation
**Rationale**: Built-in with Python, cross-platform, no additional dependencies
**Benefits**: Zero additional dependencies, professional widgets, threading support
**Trade-offs**: Not as modern as web-based UI but more reliable deployment

### 4. Executable Distribution Strategy
**Decision**: PyInstaller for standalone executable creation
**Rationale**: Enable deployment to users without Python installation
**Benefits**: Professional deployment, enterprise-ready, simplified distribution
**Implementation**: Automated build system with comprehensive dependency bundling

### 5. Project Focus Streamlining (Maintained)
**Decision**: Remove PDF processing to focus on Office document excellence
**Rationale**: PDF complexity caused installation issues and user friction
**Benefits**: 70% faster installation, eliminated problematic dependencies, improved reliability
**Implementation**: Removed AdvancedPDFProcessor class and all PDF-related dependencies

### 6. Enhanced Language Support Strategy (Maintained)
**Decision**: Expand from 3 to 6 languages with professional Asian language focus
**Languages Added**: Thai (th), Chinese Simplified (zh), Korean (ko)
**Implementation**:
```python
language_map = {
    "ja": "to Japanese",
    "vi": "to Vietnamese", 
    "en": "to English",
    "th": "to Thai",
    "zh": "to Chinese (Simplified)",
    "ko": "to Korean"
}
```

## Design Patterns

### 1. System Prompt Configuration Pattern - NEW
**Usage**: External file-based AI model instruction management
**Benefits**: Flexible behavior modification, version control, easy testing
**Implementation**: 
```python
def translate_batch():
    prompt_file = os.path.join(script_dir, "translator-system-prompt.txt")
    if os.path.exists(prompt_file):
        with open(prompt_file, 'r', encoding='utf-8') as f:
            system_prompt = f.read()
    # Use system_prompt in API call
```

### 2. Dual Interface Facade Pattern - NEW
**Usage**: Single backend serving multiple user interface types
**Benefits**: Code reuse, consistent behavior, user choice
**Implementation**: Both CLI and GUI interfaces call same translator functions

### 3. GUI Threading Pattern - NEW
**Usage**: Background processing to maintain responsive user interface
**Benefits**: Non-blocking operations, real-time progress updates
**Implementation**: 
```python
def start_translation(self):
    thread = threading.Thread(target=self.run_translation, daemon=True)
    thread.start()
```

### 4. Enhanced Strategy Pattern (Maintained)
**Usage**: Office document type processing and enhanced translation engines
**Benefits**: Easy extensibility for new Office formats and languages
**Example**: Document processors for Excel/Word/PowerPoint with six-language support

### 5. Maintained Template Method Pattern
**Usage**: Common processing workflow across Office document types
**Steps**: Load → Extract → Translate → Apply → Save
**Customization**: Each document type implements specific extraction logic
**Enhancement**: Improved with six-language support and system prompt integration

### 6. Preserved Chain of Responsibility
**Usage**: PowerPoint element extraction with multiple engines (maintained excellence)
**Flow**: Advanced extraction → COM automation → XML parsing → Basic fallback
**Benefits**: Comprehensive coverage with automatic fallback
**Status**: Fully maintained for PowerPoint advanced processing

### 7. Build Automation Pattern - NEW
**Usage**: Automated executable creation with PyInstaller
**Benefits**: Consistent builds, comprehensive dependency inclusion
**Implementation**: Custom .spec file with hidden imports and data files

## Component Relationships

### System Prompt Integration
```python
def translate_batch(texts, target_lang, api_client, progress_callback):
    # Load system prompt from external file
    system_prompt = load_system_prompt()
    
    # Use in API call
    response = api_client.chat.completions.create(
        model="gemini-2.0-flash",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
    )
```

### GUI Architecture Components
```python
class OfficeTranslatorGUI:
    # Main application window with threading
    def __init__(self, root):
        self.setup_window()
        self.setup_variables()
        self.setup_ui()
        
    def start_translation(self):
        # Background threading for responsiveness
        thread = threading.Thread(target=self.run_translation, daemon=True)
        thread.start()

class APIKeyDialog:
    # Modal dialog for API key configuration
    def __init__(self, parent, callback=None):
        # Web browser integration for key acquisition
        # Secure .env file storage
```

### Enhanced Translation Engine Core (Updated)
```python
def translate_batch(texts, target_lang, api_client, progress_callback):
    # Enhanced centralized translation logic for six languages
    # Dynamic system prompt loading
    # Used by all Office document processors and both interfaces
    # Handles batching, retries, and progress
    # Support: ja, vi, en, th, zh, ko
```

### Streamlined Document Processor Interface (Maintained)
```python
def process_*_file(input_path, target_lang, api_client, progress):
    # Common interface for Office document types only
    # Standardized error handling and progress reporting
    # Enhanced six-language support with system prompt integration
    # Returns: Optional[str] (output path or None)
```

### Build System Components - NEW
```python
# PyInstaller configuration
def create_spec_file():
    # Custom .spec with hidden imports
    # Data file inclusion (translator.py, system prompt)
    # Professional executable metadata

def build_executable():
    # Automated PyInstaller execution
    # Comprehensive dependency bundling
    # Distribution package creation
```

## Critical Implementation Paths

### 1. Excel Processing Path (Enhanced)
**Library**: xlwings (COM-based)
**Strategy**: Direct manipulation preserving formulas and formatting
**Challenges**: Shape text extraction, chart elements
**Solution**: Multi-method text extraction with fallbacks
**Enhancement**: Six-language support with system prompt integration

### 2. Word Processing Path (Enhanced)
**Library**: python-docx
**Strategy**: Document object model manipulation
**Coverage**: Paragraphs, tables, headers, footers, sections
**Preservation**: Complete formatting and structure retention
**Enhancement**: Enhanced language mapping and system prompt integration

### 3. PowerPoint Processing Path (Advanced - Maintained)
**Primary**: AdvancedPowerPointProcessor multi-engine approach (fully preserved)
**Fallback**: Basic python-pptx processing
**Advanced Elements**: SmartArt, WordArt, OLE objects, complex shapes
**Platform Features**: COM automation on Windows, XML parsing everywhere
**Status**: All advanced PowerPoint capabilities maintained and enhanced

### 4. GUI Interface Path - NEW
**Framework**: tkinter with professional widgets
**Architecture**: Threading for background processing
**Features**: Drag-drop, progress tracking, API key management
**Integration**: Direct calls to existing translation engine
**User Experience**: Modern interface matching commercial software

### 5. System Prompt Path - NEW
**Loading**: Dynamic file reading with fallback creation
**Integration**: API system message for translation instructions
**Customization**: External file modification for behavior changes
**Languages**: Six-language support with professional guidelines
**Benefits**: Flexible, maintainable, version-controlled translation behavior

### 6. Build Process Path - NEW
**Tool**: PyInstaller with custom configuration
**Dependencies**: Comprehensive Office library bundling
**Distribution**: Complete user-ready package with documentation
**Deployment**: Single executable requiring no Python installation
**Professional**: Version information and metadata inclusion

## Data Flow Patterns

### Dual Interface Processing
1. **Interface Selection**: User chooses CLI (batch file) or GUI (tkinter)
2. **Input Processing**: Both interfaces handle file discovery and validation
3. **Translation Engine**: Shared backend processes all translation requests
4. **System Prompt Integration**: Dynamic loading of translation instructions
5. **Output Management**: Consistent file organization across both interfaces

### GUI-Specific Flow
1. **User Interaction**: Point-and-click file selection and configuration
2. **Background Processing**: Threading prevents UI freezing during translation
3. **Real-time Feedback**: Progress bars and activity logging
4. **Error Handling**: User-friendly dialogs with recovery options
5. **Integration**: Seamless connection to existing translation engine

### Build and Distribution Flow
1. **Source Preparation**: All dependencies and data files identified
2. **PyInstaller Configuration**: Custom .spec file with comprehensive imports
3. **Executable Creation**: Single-file bundle with all dependencies
4. **Package Assembly**: Complete distribution folder with documentation
5. **Distribution**: Zero-dependency deployment ready for end users

## Extensibility Points

### Adding New Interfaces
1. Create new interface module following shared backend pattern
2. Implement calls to existing translation functions
3. Add interface-specific features (CLI colors, GUI widgets, etc.)
4. Maintain consistent user experience across interfaces

### Enhancing System Prompt Architecture
- **Multi-language prompts**: Separate prompt files per language
- **Domain-specific prompts**: Specialized prompts for technical/business content
- **A/B testing framework**: Easy switching between prompt versions
- **Dynamic prompt generation**: AI-generated prompts based on content analysis

### GUI Enhancement Opportunities
- **Advanced widgets**: File trees, preview panes, batch processing views
- **Themes and customization**: User-selectable color schemes and layouts
- **Integration features**: Direct cloud storage, email integration
- **Enterprise features**: User management, audit logging, batch scheduling

### Distribution Enhancements
- **Auto-update system**: Automatic updates for deployed executables
- **Enterprise deployment**: MSI packages, group policy support
- **Cloud deployment**: Web-based version with same backend
- **Mobile companion**: Mobile app connecting to desktop version

## Architecture Benefits

### Multi-Interface Strategy Benefits
- **User Choice**: Technical users can use CLI, non-technical users get GUI
- **Professional Deployment**: Executable version for enterprise distribution
- **Consistent Backend**: Same translation quality regardless of interface
- **Maintenance Efficiency**: Single codebase serving multiple interfaces

### System Prompt Architecture Benefits
- **Flexibility**: Easy behavior modification without code changes
- **Maintainability**: Centralized translation instructions
- **Testing**: Easy A/B testing of different prompt strategies
- **Version Control**: Track prompt evolution separately from code
- **Customization**: Users can modify translation behavior for specific needs

### Build System Benefits
- **Professional Deployment**: Enterprise-ready executable distribution
- **Zero Dependencies**: No Python installation required for end users
- **Simplified Support**: Fewer installation issues and support requests
- **Market Reach**: Accessible to users without technical expertise
