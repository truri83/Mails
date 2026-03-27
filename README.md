# Mails

Email management and processing application with automated module contradiction checking.

## Features

✅ **Automated Module Checking** - Automatically detect contradictions and inconsistencies in code modules
✅ **Email Processing** - Core functionality for email handling and validation
✅ **Clean Architecture** - Well-organized module structure with clear separation of concerns
✅ **CI/CD Integration** - Automated checks on every push and pull request

## Quick Start

### Check All Modules for Contradictions

Run the module contradiction checker:

```bash
python3 check_modules.py
```

This will scan all modules in the project and check for:
- Duplicate function/class definitions
- Inconsistent naming conventions
- Circular dependencies
- Conflicting constants
- Other code contradictions

### Project Structure

```
Mails/
├── src/
│   ├── core/           # Core functionality and configuration
│   ├── utils/          # Utility functions
│   ├── models/         # Data models
│   ├── services/       # Business logic services
│   └── test_helper/    # Testing utilities (modTestUnit.py)
├── tests/              # Test files
├── check_modules.py    # Module contradiction checker
└── MODULE_STANDARDS.md # Documentation for module standards
```

## Module Standards

See [MODULE_STANDARDS.md](MODULE_STANDARDS.md) for detailed information about:
- Coding conventions
- Naming standards
- Documentation requirements
- Import organization
- Best practices

## What Gets Checked

The module checker automatically detects:

### 🔴 Errors (Must Fix)
- Duplicate class definitions (same class in multiple files)
- Constants with same name but different values
- Critical naming violations

### 🟡 Warnings (Should Review)
- Duplicate function definitions
- Mixed naming conventions (snake_case vs camelCase)
- Potential circular dependencies

### ✅ Allowed Patterns
- Test mocks and implementations with same name
- Common methods like `__init__`, `__str__`, `__repr__`
- Helper functions in different modules

## CI/CD Integration

The module checker runs automatically on every push via GitHub Actions. See `.github/workflows/check-modules.yml`.

## Development

### Running Tests

```bash
# Run module checker
python3 check_modules.py

# Run on specific directory
python3 check_modules.py /path/to/modules
```

### Adding New Modules

1. Create your module in the appropriate `src/` subdirectory
2. Follow the naming conventions in MODULE_STANDARDS.md
3. Add proper documentation (docstrings)
4. Run `python3 check_modules.py` to verify no contradictions
5. Commit and push

## Example Usage

```python
from src.models.email import Email
from src.services.email_service import EmailService
from src.utils.email_utils import validate_email

# Create an email
email = Email(
    sender="sender@example.com",
    recipients=["recipient@example.com"],
    subject="Hello",
    body="This is a test email."
)

# Validate email address
if validate_email(email.sender):
    # Send email
    service = EmailService("smtp.example.com", 587)
    service.send_email(email)
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run `python3 check_modules.py` to ensure no contradictions
5. Submit a pull request

## License

This project is part of the Mails application framework.

---

**Hinweis:** Dieses Projekt prüft automatisch alle Module auf Widersprüche und Inkonsistenzen.

**Note:** This project automatically checks all modules for contradictions and inconsistencies.
