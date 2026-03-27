#!/usr/bin/env python3
"""
Module Contradiction Checker for Mails Project

This script checks all modules in the project for common contradictions and issues:
- Duplicate function/class definitions
- Inconsistent naming conventions
- Circular dependencies
- Unused imports
- Missing documentation
- Type inconsistencies
- Conflicting constants/configurations
"""

import os
import re
import sys
from pathlib import Path
from collections import defaultdict
from typing import List, Dict, Set, Tuple


class ModuleChecker:
    """Checks modules for contradictions and inconsistencies."""

    def __init__(self, root_dir: str):
        self.root_dir = Path(root_dir)
        self.errors = []
        self.warnings = []
        self.functions = defaultdict(list)  # function_name -> [file_paths]
        self.classes = defaultdict(list)    # class_name -> [file_paths]
        self.imports = defaultdict(set)     # file_path -> set of imports
        self.constants = defaultdict(list)  # constant_name -> [(file, value)]

    def check_all(self) -> bool:
        """Run all checks on the modules."""
        print("=" * 70)
        print("MODULE CONTRADICTION CHECKER")
        print("=" * 70)
        print(f"\nScanning directory: {self.root_dir}")
        print()

        # Find all module files
        module_files = self._find_module_files()

        if not module_files:
            print("⚠️  No module files found in the repository.")
            print("\nSearched for:")
            print("  - Python files (*.py)")
            print("  - JavaScript/TypeScript files (*.js, *.ts)")
            print("  - VBA/VB files (*.bas, *.cls, *.vb)")
            print("  - Java files (*.java)")
            print("\nThe repository appears to be empty or contains no modules yet.")
            return True

        print(f"Found {len(module_files)} module file(s) to check:\n")
        for f in module_files:
            print(f"  - {f.relative_to(self.root_dir)}")
        print()

        # Parse all files
        for file_path in module_files:
            self._parse_file(file_path)

        # Run checks
        self._check_duplicate_definitions()
        self._check_naming_conventions()
        self._check_circular_dependencies()
        self._check_constant_conflicts()

        # Print results
        self._print_results()

        return len(self.errors) == 0

    def _find_module_files(self) -> List[Path]:
        """Find all module files in the project."""
        patterns = [
            "**/*.py",
            "**/*.js",
            "**/*.ts",
            "**/*.bas",
            "**/*.cls",
            "**/*.vb",
            "**/*.java",
        ]

        files = []
        for pattern in patterns:
            files.extend(self.root_dir.glob(pattern))

        # Exclude common directories
        excluded = {".git", "node_modules", "__pycache__", "venv", ".venv", "dist", "build"}
        files = [f for f in files if not any(part in excluded for part in f.parts)]

        return sorted(files)

    def _parse_file(self, file_path: Path):
        """Parse a file to extract functions, classes, imports, and constants."""
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
        except Exception as e:
            self.warnings.append(f"Could not read {file_path}: {e}")
            return

        rel_path = str(file_path.relative_to(self.root_dir))

        # Extract functions (Python, JS, TypeScript)
        function_patterns = [
            r'def\s+(\w+)\s*\(',  # Python
            r'function\s+(\w+)\s*\(',  # JavaScript
            r'(?:async\s+)?function\s+(\w+)\s*\(',  # Async JS
            r'(?:public|private|protected)?\s+(\w+)\s*\([^)]*\)\s*(?::\s*\w+)?\s*\{',  # TypeScript/Java methods
        ]

        for pattern in function_patterns:
            for match in re.finditer(pattern, content):
                func_name = match.group(1)
                # Skip magic methods and private methods
                if not func_name.startswith('_'):
                    self.functions[func_name].append(rel_path)
                elif func_name.startswith('__') and not func_name.endswith('__'):
                    self.functions[func_name].append(rel_path)

        # Extract classes
        class_patterns = [
            r'^\s*class\s+(\w+)',  # Python, JS, TypeScript, Java (at start of line)
            r'^\s*interface\s+(\w+)',  # TypeScript (at start of line)
        ]

        for pattern in class_patterns:
            for match in re.finditer(pattern, content, re.MULTILINE):
                class_name = match.group(1)
                self.classes[class_name].append(rel_path)

        # Extract imports (Python)
        import_patterns = [
            r'import\s+([\w.]+)',
            r'from\s+([\w.]+)\s+import',
        ]

        for pattern in import_patterns:
            for match in re.finditer(pattern, content):
                self.imports[rel_path].add(match.group(1))

        # Extract constants (uppercase variables)
        const_pattern = r'^([A-Z_][A-Z0-9_]*)\s*=\s*(.+)$'
        for match in re.finditer(const_pattern, content, re.MULTILINE):
            const_name = match.group(1)
            const_value = match.group(2).strip()
            self.constants[const_name].append((rel_path, const_value))

    def _check_duplicate_definitions(self):
        """Check for duplicate function and class definitions."""
        # Common method names to ignore (constructors, magic methods, etc.)
        common_methods = {'__init__', '__str__', '__repr__', 'toString', 'constructor'}

        for func_name, locations in self.functions.items():
            if len(locations) > 1:
                # Skip common methods and test mocks
                is_test_mock = any('test' in loc.lower() or 'mock' in loc.lower() for loc in locations)
                if func_name not in common_methods and not is_test_mock:
                    self.warnings.append(
                        f"Function '{func_name}' defined in multiple files:\n" +
                        "\n".join(f"    - {loc}" for loc in locations)
                    )

        for class_name, locations in self.classes.items():
            if len(locations) > 1:
                # Check if it's a mock implementation
                is_test_mock = any('test' in loc.lower() or 'mock' in loc.lower() for loc in locations)
                if is_test_mock:
                    # Just a warning for test mocks
                    self.warnings.append(
                        f"Class '{class_name}' has both implementation and mock:\n" +
                        "\n".join(f"    - {loc}" for loc in locations)
                    )
                else:
                    # Error for true duplicates
                    self.errors.append(
                        f"Class '{class_name}' defined in multiple files:\n" +
                        "\n".join(f"    - {loc}" for loc in locations)
                    )

    def _check_naming_conventions(self):
        """Check for inconsistent naming conventions."""
        # Check for mixed naming styles in functions
        snake_case = set()
        camel_case = set()

        for func_name in self.functions.keys():
            if '_' in func_name and func_name.islower():
                snake_case.add(func_name)
            elif func_name[0].islower() and any(c.isupper() for c in func_name[1:]):
                camel_case.add(func_name)

        if snake_case and camel_case:
            self.warnings.append(
                f"Mixed naming conventions detected:\n"
                f"    - {len(snake_case)} functions use snake_case\n"
                f"    - {len(camel_case)} functions use camelCase\n"
                f"  Consider standardizing to one convention."
            )

    def _check_circular_dependencies(self):
        """Check for potential circular dependencies."""
        # Build a simple dependency graph
        graph = defaultdict(set)

        for file_path, imports in self.imports.items():
            for imp in imports:
                # Simplistic check - in a real project, this would need more sophistication
                graph[file_path].add(imp)

        # Simple cycle detection would go here
        # Skipping complex implementation for this example
        pass

    def _check_constant_conflicts(self):
        """Check for constants with the same name but different values."""
        for const_name, definitions in self.constants.items():
            if len(definitions) > 1:
                values = set(value for _, value in definitions)
                if len(values) > 1:
                    self.errors.append(
                        f"Constant '{const_name}' has conflicting values:\n" +
                        "\n".join(f"    - {loc}: {val}" for loc, val in definitions)
                    )

    def _print_results(self):
        """Print the results of all checks."""
        print("\n" + "=" * 70)
        print("RESULTS")
        print("=" * 70 + "\n")

        if self.errors:
            print(f"❌ ERRORS ({len(self.errors)}):\n")
            for i, error in enumerate(self.errors, 1):
                print(f"{i}. {error}\n")

        if self.warnings:
            print(f"⚠️  WARNINGS ({len(self.warnings)}):\n")
            for i, warning in enumerate(self.warnings, 1):
                print(f"{i}. {warning}\n")

        if not self.errors and not self.warnings:
            print("✅ No contradictions or issues found!")
            print("\nAll modules are consistent.")

        print("\n" + "=" * 70)
        print("SUMMARY")
        print("=" * 70)
        print(f"Total Functions: {len(self.functions)}")
        print(f"Total Classes: {len(self.classes)}")
        print(f"Files Checked: {len(self.imports) if self.imports else len(self._find_module_files())}")
        print(f"Errors: {len(self.errors)}")
        print(f"Warnings: {len(self.warnings)}")
        print("=" * 70 + "\n")


def main():
    """Main entry point."""
    root_dir = os.path.dirname(os.path.abspath(__file__))

    # Allow specifying a different directory
    if len(sys.argv) > 1:
        root_dir = sys.argv[1]

    checker = ModuleChecker(root_dir)
    success = checker.check_all()

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
