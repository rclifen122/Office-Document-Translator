# 🤝 Contributing to Office Document Translator

We love your input! We want to make contributing to Office Document Translator as easy and transparent as possible, whether it's:

- 🐛 Reporting a bug
- 💡 Discussing the current state of the code
- 🚀 Submitting a fix
- 💪 Proposing new features
- 👨‍💻 Becoming a maintainer

## 📋 Development Process

We use GitHub to host code, to track issues and feature requests, as well as accept pull requests.

### 🔄 Pull Request Process

1. **Fork** the repository and create your branch from `main`
2. **Add tests** if you've added code that should be tested
3. **Update documentation** if you've changed APIs
4. **Ensure** the test suite passes
5. **Make sure** your code follows the existing style
6. **Submit** that pull request!

## 🐛 Bug Reports

We use GitHub issues to track public bugs. Report a bug by [opening a new issue](https://github.com/rclifen122/Office-Document-Translator/issues/new/choose).

**Great Bug Reports** tend to have:

- 📝 A quick summary and/or background
- 🔍 Specific steps to reproduce
  - Be specific!
  - Give sample code if you can
- 🎯 What you expected would happen
- 😵 What actually happens
- 💭 Additional context
  - Include error messages, logs, screenshots
  - System information (OS, Python version)

## 🚀 Feature Requests

We welcome feature requests! Please provide:

- 🎯 **Clear description** of the feature
- 💡 **Use case** - why do you need this feature?
- 📝 **Example** of how it would work
- 🤔 **Alternatives** you've considered

## 💻 Development Setup

1. **Fork and clone** the repository
   ```bash
   git clone https://github.com/your-username/Office-Document-Translator.git
   cd Office-Document-Translator
   ```

2. **Install dependencies**
   ```bash
   pip install -r translator-requirements.txt
   pip install -r requirements_exe.txt  # For building executables
   ```

3. **Set up environment**
   ```bash
   cp .env.example .env  # Create from template
   # Edit .env with your API keys
   ```

4. **Run tests**
   ```bash
   python -m pytest tests/
   ```

## 📐 Code Style

- **Follow PEP 8** Python style guidelines
- **Use meaningful** variable and function names
- **Add docstrings** to functions and classes
- **Comment complex** logic
- **Keep functions** focused and small

### Example Code Style:
```python
def translate_document(file_path: str, target_language: str) -> bool:
    """
    Translate a document to the target language.
    
    Args:
        file_path (str): Path to the document to translate
        target_language (str): Target language code (ja, en, vi)
        
    Returns:
        bool: True if translation successful, False otherwise
    """
    # Implementation here
    pass
```

## 🧪 Testing

- **Write tests** for new features
- **Ensure** existing tests pass
- **Use descriptive** test names
- **Test edge cases** and error conditions

## 📚 Documentation

- **Update README.md** for user-facing changes
- **Add docstrings** to new functions/classes
- **Update inline comments** as needed
- **Create examples** for new features

## 🏷️ Commit Message Guidelines

Use clear and meaningful commit messages:

```
feat: add support for PowerPoint animations
fix: resolve Excel cell formatting issue
docs: update installation instructions
test: add tests for Word document processing
refactor: improve error handling in translator
```

### Types:
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `test`: Adding or modifying tests
- `refactor`: Code refactoring
- `style`: Code style changes
- `perf`: Performance improvements

## 🏅 Recognition

Contributors will be acknowledged in:
- 📄 README.md contributors section
- 🏆 GitHub releases
- 💬 Project discussions

## 📄 License

By contributing, you agree that your contributions will be licensed under the MIT License.

## ❓ Questions?

Feel free to:
- 💬 Open a [discussion](https://github.com/rclifen122/Office-Document-Translator/discussions)
- 📧 Create an [issue](https://github.com/rclifen122/Office-Document-Translator/issues)
- 📧 Contact the maintainers

---

Thank you for contributing! 🎉 