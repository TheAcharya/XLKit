# XLKit Test Suite Documentation

This document provides an organized overview of all tests in the XLKit library, grouped by feature area. Each section describes the purpose and key features of the tests, with no repeated or redundant information.

## Table of Contents
- [Test Overview](#test-overview)
- [Core Workbook & Sheet Tests](#core-workbook--sheet-tests)
- [Cell Operations & Data Types](#cell-operations--data-types)
- [Coordinate & Range Tests](#coordinate--range-tests)
- [File Operations](#file-operations)
- [Image & Aspect Ratio Tests](#image--aspect-ratio-tests)
- [Cell Formatting](#cell-formatting)
- [CSV/TSV Import/Export](#csvtsv-importexport)
- [Column & Row Sizing](#column--row-sizing)
- [XLKitTestRunner](#xlkittestrunner)
- [Test Execution & Validation](#test-execution--validation)
- [Coverage & Quality Assurance](#coverage--quality-assurance)

## Test Overview
- Total Tests: 46
- 100% coverage of public APIs
- All generated files validated with CoreXLSX
- Security features integrated throughout all tests

## Core Workbook & Sheet Tests
- `testCreateWorkbook()`: Basic workbook creation and initialization
- `testAddSheet()`: Sheet addition with proper ID assignment
- `testGetSheetByName()`: Sheet retrieval by name functionality
- `testRemoveSheet()`: Sheet removal and count tracking
- `testComplexWorkbook()`: Complex workbook with multiple sheets and formulas

## Cell Operations & Data Types
- `testSetAndGetCell()`: Setting/getting all cell value types (string, number, integer, boolean, date, formula)
- `testSetCellByRowColumn()`: Cell operations using row/column coordinates
- `testSetRange()`: Range-based cell operations
- `testMergeCells()`: Cell merging functionality
- `testGetUsedCells()`: Used cell detection
- `testClearSheet()`: Sheet clearing operations
- `testCellValueStringValue()`: String representation for all cell values
- `testCellValueType()`: Type identification for all cell values

## Coordinate & Range Tests
- `testCellCoordinate()`: Excel coordinate creation and parsing
- `testCellRange()`: Range creation, notation parsing, and coordinate enumeration

## File Operations
- `testSaveWorkbook()`: Async workbook saving and file validation
- `testSaveWorkbookSync()`: Synchronous workbook saving and file validation

## Image & Aspect Ratio Tests
- `testImageFormatDetection()`: Image format detection (GIF, PNG, JPEG)
- `testImageSizeDetection()`: Image size detection and metadata
- `testExcelImageCreation()`: ExcelImage object creation and validation
- `testSheetImageOperations()`: Sheet-level image operations
- `testWorkbookImageOperations()`: Workbook-level image operations
- `testGIFEmbedding()`: GIF image embedding with validation
- `testMultipleImageFormats()`: Multi-format image support
- `testSimplifiedImageEmbeddingAPI()`: Easy-to-use image embedding API
- `testAspectRatioPreservation()`: Basic aspect ratio preservation testing
- `testAspectRatioPreservationWithDifferentSizes()`: Aspect ratio preservation with various image sizes
- `testAspectRatioPreservationForAnyDimensions()`: Comprehensive aspect ratio testing for all 17 professional ratios
- `testExcelCellDimensionsForAnyImageSize()`: Excel cell dimension calculations for any image size

### Aspect Ratio Preservation Testing
XLKit extensively tests all 17 professional video and cinema aspect ratios:
- 16:9 (HD/4K video)
- 1:1 (Square format)
- 9:16 (Vertical video)
- 21:9 (Ultra-wide)
- 3:4 (Portrait)
- 2.39:1 (Cinemascope/Anamorphic)
- 1.85:1 (Academy ratio)
- 4:3 (Classic TV/monitor)
- 18:9 (Modern mobile)
- 1.19:1 (HD Standard)
- 1.5:1 (SD Academy)
- 1.48:1 (SD Academy Alt)
- 1.25:1 (SD Standard)
- 1.9:1 (IMAX Digital)
- 1.32:1 (DCI Standard)
- 2.37:1 (5K Cinema Scope)
- 1.37:1 (IMAX Film 15/70mm)

All tests verify pixel-perfect scaling, Excel cell dimension matching, and zero distortion.

## Cell Formatting
- `testCellFormatting()`: Basic cell formatting with predefined styles
- `testCustomCellFormatting()`: Custom formatting with full property customization
- `testRangeFormatting()`: Range-based formatting and persistence

## CSV/TSV Import/Export
- `testCSVExport()`: CSV export functionality
- `testTSVExport()`: TSV export functionality
- `testCSVImport()`: CSV import with header handling
- `testTSVImport()`: TSV import with delimiter recognition
- `testCSVImportWithQuotes()`: CSV import with quoted values
- `testCSVImportWithDates()`: CSV import with date handling
- `testCSVImportIntoExistingSheet()`: Import into existing sheets
- `testCSVExportWithSpecialCharacters()`: Export with special character handling

## Column & Row Sizing
- `testColumnWidthOperations()`: Column width management and persistence
- `testRowHeightOperations()`: Row height management and persistence
- `testAutoSizeColumnForImage()`: Automatic column sizing for images

## XLKitTestRunner

A modular command-line tool for generating Excel files for testing and demonstration purposes.

### Available Test Types
- `no-embeds` / `no-images`: Generate Excel from CSV without images
- `embed` / `with-embeds` / `with-images`: Generate Excel with embedded images from CSV data
- `comprehensive` / `demo`: Comprehensive API demonstration with all features
- `security-demo` / `security`: Demonstrate file path security restrictions
- `help` / `-h` / `--help`: Show available commands

### Test Features
- Security Integration: All tests include security logging and validation
- CoreXLSX Validation: Generated files are validated for Excel compliance
- Aspect Ratio Testing: Image embedding tests all 17 professional aspect ratios
- Performance Testing: Large dataset handling and memory optimization
- Error Handling: Comprehensive error testing and edge case coverage

### Output Structure
```
Test-Workflows/
├── Embed-Test.xlsx          # From no-embeds test
├── Embed-Test-Embed.xlsx    # From embed test (with images)
├── Comprehensive-Demo.xlsx  # From comprehensive test
└── [Your-Test].xlsx         # From custom tests
```

### Security Features in Tests
- Rate Limiting: Prevents test abuse and resource exhaustion
- Security Logging: All test operations are logged for audit trails
- Input Validation: All test inputs are validated for security
- File Quarantine: Suspicious test files are automatically quarantined
- Checksum Verification: Optional file integrity verification (disabled by default)

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

### Running XLKitTestRunner
```bash
# Run specific test types
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner help
```

### Validation
- All tests generate valid Excel files
- CoreXLSX validation for Excel compliance
- Aspect ratio tests verify pixel-perfect preservation
- File operations tests validate complete workflows
- Security features are validated in all operations

## Coverage & Quality Assurance

| Category           | Test Count | Coverage         |
|--------------------|------------|-----------------|
| Core Workbook      | 5          | 100%            |
| Sheet Management   | 5          | 100%            |
| Cell Operations    | 8          | 100%            |
| Data Types         | 2          | 100%            |
| Coordinates/Ranges | 2          | 100%            |
| File Operations    | 2          | 100%            |
| Image Support      | 12         | 100%            |
| Aspect Ratios      | 3          | 100% (17 ratios)|
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
- Security features integrated throughout all tests
- XLKitTestRunner provides comprehensive demonstration capabilities

### Security Integration
- Rate limiting prevents test abuse
- Security logging captures all operations
- Input validation ensures data integrity
- File quarantine protects against malicious content
- Checksum verification ensures file authenticity

This suite ensures XLKit delivers reliable, high-quality Excel file generation with perfect image embedding and aspect ratio preservation, backed by comprehensive security features and validation. 