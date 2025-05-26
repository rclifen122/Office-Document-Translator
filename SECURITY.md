# 🔒 Security Policy

## 🛡️ Supported Versions

We release patches for security vulnerabilities for the following versions:

| Version | Supported          |
| ------- | ------------------ |
| 2.0.x   | ✅ |
| 1.0.x   | ❌ |

## 🚨 Reporting a Vulnerability

We take the security of Office Document Translator seriously. If you believe you have found a security vulnerability, please report it to us as described below.

### ⚠️ Please DO NOT report security vulnerabilities through public GitHub issues.

Instead, please report them via:

1. **Email**: [Create a private issue or email the maintainer]
2. **GitHub Security Advisories**: Use the [security advisory feature](https://github.com/rclifen122/Office-Document-Translator/security/advisories/new)

### 📝 What to Include

Please include the following information in your report:

- **Description** of the vulnerability
- **Steps to reproduce** the issue
- **Potential impact** of the vulnerability
- **Suggested fix** (if you have one)
- **Your contact information** for follow-up questions

### 🔄 Response Timeline

- **Initial Response**: Within 48 hours
- **Investigation**: Within 1 week
- **Fix Release**: Depending on severity, within 2-4 weeks
- **Public Disclosure**: After fix is released

## 🛡️ Security Considerations

### 🔑 API Keys
- Never commit API keys to the repository
- Use `.env` files for local development
- Use secure secret management in production

### 📄 Document Processing
- Documents are processed locally on your machine
- No documents are uploaded to external servers (except for translation API calls)
- Temporary files are cleaned up after processing

### 🌐 Network Security
- All API communications use HTTPS
- API keys are transmitted securely
- No sensitive data is logged

### 🔐 Data Privacy
- Original documents remain on your local machine
- Only text content is sent to translation APIs
- No personal data is collected or stored

## 🏆 Recognition

We appreciate security researchers who help keep our users safe. Reporters of valid security issues will be:

- Credited in our security acknowledgments (if desired)
- Mentioned in release notes (if appropriate)
- Given priority support for their reports

## 📚 Resources

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [Python Security Guidelines](https://python.org/dev/security/)
- [Google AI Responsible AI Practices](https://ai.google/responsibilities/responsible-ai-practices/)

---

Thank you for helping keep Office Document Translator and our users safe! 🙏 