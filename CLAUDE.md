# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A utility library for Excel and PDF manipulation using Apache POI and PDFBox. Multi-module Maven project with Jakarta EE support.

- **Java**: 21
- **Build tool**: Maven
- **Main modules**: `ecuacion-util-excel-table`, `ecuacion-util-excel-report-to-pdf`, `ecuacion-util-excel-table-sample`

## Java Coding Rules

### Style Standards
- Follows **Google Java Style Guide** (enforced by Checkstyle)
- Indentation: **2 spaces** (no tabs)
- Max line length: **100 characters** (excluding package/import statements) — **applies to comments too**
- Encoding: **UTF-8**

### Imports
- Wildcard imports (`.*`) are **prohibited**
- Imports are sorted automatically (follow IDE auto-organize imports)

### Javadoc
- **All public classes, methods, and fields must have Javadoc**
- `@return` and `@param` tags must not be omitted
- When editing existing files, review and update Javadoc for any modified methods

### License Header
- All Java files must have the Apache 2.0 license header at the top
- Follow the same format as existing files

## File Creation and Editing Rules

- Always refer to existing files in the same package before creating a new one
- When adding to a package that has `package-info.java`, check its contents first

## Work Style

- **Commit only when explicitly instructed**
- **Push only when explicitly instructed**
- Always confirm before destructive operations (file deletion, `git reset --hard`, etc.)
- Do not propose changes to code that has not been read first

## Build and Verification

```bash
# Build all modules
mvn clean install

# Build a specific module
mvn clean install -pl ecuacion-util-excel-table

# Run all tests
mvn test

# Run a single test class
mvn test -Dtest=ExcelReadUtilTest

# Run a single test method
mvn test -Dtest=ExcelReadUtilTest#methodName
```

**Always run the following after editing Java files and fix any violations before finishing:**

```bash
mvn checkstyle:check spotbugs:check
mvn javadoc:javadoc
```

Common violations:
- Checkstyle: Line length over 100 characters (including comments and Javadoc)
- Checkstyle: Missing Javadoc on public members
- Checkstyle: Wildcard imports
- SpotBugs: Using reflection to access private fields (use `protected` scope workarounds where needed)

## Architecture Overview

### `ecuacion-util-excel-table` Internal Structure

```
jp.ecuacion.util.excel/
├── util/          # ExcelReadUtil, ExcelWriteUtil (low-level Cell operations)
├── table/
│   ├── reader/    # Interfaces + concrete/ for implementations
│   └── writer/    # Interfaces + concrete/ for implementations
├── enums/         # NoDataString, etc.
├── exception/
└── spi/           # MessagesUtilExcelTableProvider (i18n SPI)
```

Table reader/writer implementations live in `table/reader/concrete/` and `table/writer/concrete/`. The low-level utilities in `util/` serve as their foundation.

### Dependency Prerequisites

The parent POM references `ecuacion-lib` in the sibling directory:
```
../ecuacion-lib/ecuacion-lib-dependency-jakartaee
```
Building this project requires `ecuacion-lib` to exist at the same directory level and be installed locally.

### i18n

Message resources are located in `src/main/resources/`:

- `messages_util_excel_table.properties` (English)
- `messages_util_excel_table_ja.properties` (Japanese)
