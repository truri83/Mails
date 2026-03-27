# Module Standards and Contradiction Checking

## Übersicht / Overview

Dieses Dokument beschreibt die Standards für Module im Mails-Projekt und wie Widersprüche automatisch geprüft werden können.

This document describes the standards for modules in the Mails project and how contradictions are automatically checked.

## Automatische Prüfung / Automatic Checking

Führen Sie den Module-Checker aus mit / Run the module checker with:

```bash
python3 check_modules.py
```

oder von einem anderen Verzeichnis / or from another directory:

```bash
python3 check_modules.py /path/to/project
```

## Geprüfte Widersprüche / Checked Contradictions

Der Checker prüft auf folgende Probleme / The checker tests for the following issues:

### 1. Doppelte Definitionen / Duplicate Definitions

**Problem:** Funktionen oder Klassen mit demselben Namen in mehreren Dateien.
**Problem:** Functions or classes with the same name in multiple files.

**Lösung:** Umbenennen oder konsolidieren Sie doppelte Definitionen.
**Solution:** Rename or consolidate duplicate definitions.

### 2. Inkonsistente Namenskonventionen / Inconsistent Naming Conventions

**Problem:** Mischung aus verschiedenen Namensstilen (z.B. snake_case und camelCase).
**Problem:** Mixing different naming styles (e.g., snake_case and camelCase).

**Empfohlene Standards:**
- Python: `snake_case` für Funktionen und Variablen, `PascalCase` für Klassen
- JavaScript/TypeScript: `camelCase` für Funktionen und Variablen, `PascalCase` für Klassen
- Konstanten: `UPPER_SNAKE_CASE` in allen Sprachen

**Recommended Standards:**
- Python: `snake_case` for functions and variables, `PascalCase` for classes
- JavaScript/TypeScript: `camelCase` for functions and variables, `PascalCase` for classes
- Constants: `UPPER_SNAKE_CASE` in all languages

### 3. Konstanten-Konflikte / Constant Conflicts

**Problem:** Konstanten mit demselben Namen aber unterschiedlichen Werten.
**Problem:** Constants with the same name but different values.

**Lösung:** Verwenden Sie eine zentrale Konfigurationsdatei für gemeinsame Konstanten.
**Solution:** Use a central configuration file for shared constants.

**Beispiel / Example:**

```python
# config.py
DATABASE_HOST = "localhost"
DATABASE_PORT = 5432
MAX_RETRIES = 3
```

### 4. Zirkuläre Abhängigkeiten / Circular Dependencies

**Problem:** Modul A importiert Modul B, das wiederum Modul A importiert.
**Problem:** Module A imports Module B, which in turn imports Module A.

**Lösung:** Refaktorieren Sie den Code, um die zirkuläre Abhängigkeit zu entfernen.
**Solution:** Refactor the code to remove the circular dependency.

## Module Standards

### Dateiorganisation / File Organization

```
Mails/
├── src/
│   ├── core/           # Kernfunktionalität / Core functionality
│   ├── utils/          # Hilfsfunktionen / Utility functions
│   ├── models/         # Datenmodelle / Data models
│   ├── services/       # Dienste / Services
│   └── test_helper/    # Testhilfen / Test helpers
├── tests/              # Tests
└── docs/               # Dokumentation / Documentation
```

### Dokumentationsstandards / Documentation Standards

Jedes Modul sollte enthalten / Each module should include:

1. **Dateikopf / File Header:**
   ```python
   """
   Module Name: Brief description

   Purpose: Detailed description of the module's purpose

   Author: Name
   Date: YYYY-MM-DD
   """
   ```

2. **Funktionsdokumentation / Function Documentation:**
   ```python
   def function_name(param1, param2):
       """
       Brief description of the function.

       Args:
           param1: Description of param1
           param2: Description of param2

       Returns:
           Description of return value

       Raises:
           ExceptionType: When this exception is raised
       """
   ```

### Import-Organisation / Import Organization

Organisieren Sie Imports in dieser Reihenfolge / Organize imports in this order:

1. Standard-Bibliothek / Standard library imports
2. Drittanbieter-Imports / Third-party imports
3. Lokale Anwendungsimports / Local application imports

```python
# Standard library
import os
import sys

# Third-party
import requests
import numpy as np

# Local application
from src.core import database
from src.utils import helpers
```

## Best Practices

### 1. Einheitliche Fehlerbehandlung / Consistent Error Handling

```python
class MailsError(Exception):
    """Base exception for Mails project."""
    pass

class DatabaseError(MailsError):
    """Database-related errors."""
    pass
```

### 2. Konfigurationsmanagement / Configuration Management

Verwenden Sie Umgebungsvariablen für konfigurierbare Werte:
Use environment variables for configurable values:

```python
import os

DATABASE_URL = os.getenv('DATABASE_URL', 'postgresql://localhost/mails')
DEBUG = os.getenv('DEBUG', 'False').lower() == 'true'
```

### 3. Logging

```python
import logging

logger = logging.getLogger(__name__)

def process_mail(mail_id):
    logger.info(f"Processing mail {mail_id}")
    try:
        # Process
        logger.debug(f"Mail {mail_id} processed successfully")
    except Exception as e:
        logger.error(f"Error processing mail {mail_id}: {e}")
        raise
```

## Continuous Integration

Der Module-Checker sollte als Teil des CI/CD-Prozesses ausgeführt werden:
The module checker should be run as part of the CI/CD process:

```yaml
# .github/workflows/check-modules.yml
name: Check Modules

on: [push, pull_request]

jobs:
  check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - name: Check modules for contradictions
        run: python3 check_modules.py
```

## Erweiterung des Checkers / Extending the Checker

Um neue Prüfungen hinzuzufügen, erweitern Sie die `ModuleChecker`-Klasse:
To add new checks, extend the `ModuleChecker` class:

```python
def _check_custom_rule(self):
    """Check for custom project-specific rules."""
    # Implementation
    pass
```

## Fragen? / Questions?

Bei Fragen zu den Modulstandards oder dem Checker, erstellen Sie ein Issue im Repository.
For questions about module standards or the checker, create an issue in the repository.
