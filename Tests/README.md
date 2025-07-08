# XLKit Test Suite Documentation

This document provides an organized overview of all tests in the XLKit library, grouped by feature area. Each section describes the purpose and key features of the tests, with no repeated or redundant information.

## Table of Contents
- [Test Overview](#test-overview)
- [Core Workbook & Sheet Tests](#core-workbook--sheet-tests)
- [Cell Operations & Data Types](#cell-operations--data-types)
- [Coordinate & Range Tests](#coordinate--range-tests)
- [Utility & File Operations](#utility--file-operations)
- [Image & Aspect Ratio Tests](#image--aspect-ratio-tests)
- [Cell Formatting](#cell-formatting)
- [CSV/TSV Import/Export](#csvtsv-importexport)
- [Column & Row Sizing](#column--row-sizing)
- [Test Execution & Validation](#test-execution--validation)
- [Coverage & Quality Assurance](#coverage--quality-assurance)

## Test Overview
- Total Tests: 46
- 100% coverage of public APIs
- All generated files validated with CoreXLSX

## Core Workbook & Sheet Tests
- Workbook creation, complex workbook with multiple sheets and formulas
- Sheet addition, retrieval, removal, and count tracking

## Cell Operations & Data Types
- Setting/getting all cell value types (string, number, integer, boolean, date, formula)
- Range-based cell operations and merging
- Used cell detection and sheet clearing
- String representation and type identification for all cell values

## Coordinate & Range Tests
- Excel coordinate creation and parsing
- Range creation, notation parsing, and coordinate enumeration

## Utility & File Operations
- Column letter/number conversion
- Date conversion utilities
- XML string escaping
- Async and sync workbook saving, file existence validation

## Image & Aspect Ratio Tests
- Image format and size detection (GIF, PNG, JPEG)
- ExcelImage object creation and metadata
- Sheet and workbook image operations
- GIF, PNG, JPEG embedding and multi-format support
- **Aspect Ratio Preservation:**
  - Tests 17 professional and common aspect ratios:
    - 16:9, 1:1, 9:16, 21:9, 3:4, 2.39:1, 1.85:1, 4:3, 18:9, 1.19:1, 1.5:1, 1.48:1, 1.25:1, 1.9:1, 1.32:1, 2.37:1, 1.37:1
  - Verifies pixel-perfect scaling, Excel cell dimension matching, and zero distortion
  - Automatic cell sizing and Excel compliance
- Simplified image embedding API: easy-to-use, automatic scaling, and registration

## Cell Formatting
- Basic and custom cell formatting (header, currency, percentage, date, borders)
- Full property customization: font, color, alignment, wrapping, number formats, borders
- Range-based formatting and persistence

## CSV/TSV Import/Export
- CSV/TSV export and import, including header handling, delimiter recognition, and data formatting
- Quoted value and date handling, special character processing
- Import into existing sheets and append functionality

## Column & Row Sizing
- Column width and row height management, retrieval, and persistence
- Automatic column sizing for images

## Test Execution & Validation

### Running All Tests
```bash
swift test
```

### Running Specific Test Categories
```bash
# Image-related tests
swift test --filter "test.*Image.*|test.*Aspect.*"

# CSV/TSV tests
swift test --filter "test.*CSV.*|test.*TSV.*"

# Core functionality tests
swift test --filter "testCreateWorkbook|testAddSheet|testSetAndGetCell"
```

### Validation
- All tests generate valid Excel files
- CoreXLSX validation for Excel compliance
- Aspect ratio tests verify pixel-perfect preservation
- File operations tests validate complete workflows

## Coverage & Quality Assurance

| Category           | Test Count | Coverage         |
|--------------------|------------|-----------------|
| Core Workbook      | 3          | 100%            |
| Sheet Management   | 3          | 100%            |
| Cell Operations    | 6          | 100%            |
| Data Types         | 2          | 100%            |
| Coordinates/Ranges | 2          | 100%            |
| Utilities          | 3          | 100%            |
| File Operations    | 2          | 100%            |
| Image Support      | 7          | 100%            |
| Aspect Ratios      | 5          | 100% (17 ratios)|
| Cell Formatting    | 3          | 100%            |
| CSV/TSV            | 8          | 100%            |
| Column/Row Sizing  | 3          | 100%            |
| **Total**          | **46**     | **100%**        |

### Quality Standards
- All generated files pass CoreXLSX validation
- Zero tolerance for aspect ratio distortion
- Efficient memory usage and file generation
- Comprehensive error and edge case testing
- 100% of public APIs tested
- Automated CI on macOS, performance and memory monitoring

This suite ensures XLKit delivers reliable, high-quality Excel file generation with perfect image embedding and aspect ratio preservation. 