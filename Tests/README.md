# XLKit Test Suite Documentation

This document provides an organized overview of all tests in the **`XLKitTests`** target (library unit tests), grouped by feature area. Each section describes the purpose and key features of the tests, with no repeated or redundant information.

**Related (not `swift test`):** **`XLKitTestRunner`** generates sample `.xlsx` files and provides the `sheet-password` helper — see [XLKitTestRunner](#xlkittestrunner), **`Test-Workflows/README.md`**, and **`Sources/XLKitTestRunner/README.md`**.

## Table of Contents
- [Test Overview](#test-overview)
- [Test File Structure](#test-file-structure)
- [Core Workbook & Sheet Tests](#core-workbook--sheet-tests)
- [Cell Operations & Data Types](#cell-operations--data-types)
- [Coordinate & Range Tests](#coordinate--range-tests)
- [File Operations](#file-operations)
- [Image & Aspect Ratio Tests](#image--aspect-ratio-tests)
- [Cell Formatting](#cell-formatting)
- [Border & Merge Tests](#border--merge-tests)
- [Number Format Tests](#number-format-tests)
- [Text Wrapping Tests](#text-wrapping-tests)
- [CSV/TSV Import/Export](#csvtsv-importexport)
- [Column & Row Sizing](#column--row-sizing)
- [Column Ordering Tests](#column-ordering-tests)
- [Sheet Utility Tests](#sheet-utility-tests)
- [Sheet State Tests](#sheet-state-tests)
- [Sheet Protection Tests](#sheet-protection-tests)
- [XLKitTestRunner](#xlkittestrunner)
- [Test Execution & Validation](#test-execution--validation)
- [Coverage & Quality Assurance](#coverage--quality-assurance)

## Test Overview
- Total Tests: **80** (15 test suites + shared `XLKitTestSupport` in `XLKitTestBase.swift`)
- Framework: **Swift Testing** (`import Testing`, `@Suite`, `@Test`, `#expect` / `#require`)
- 100% coverage of public APIs
- Save and file-operation tests write real `.xlsx` files via temp-workbook helpers
- Security features integrated throughout save and validation paths
- Text alignment (horizontal, vertical, combined) is fully tested
- Border and merge functionality is fully tested
- Text wrapping functionality is fully tested
- Column ordering for sheets with more than 26 columns is fully tested
- CSV/TSV edge cases (quoted fields, escaped quotes, empty fields) are fully tested
- Sheet visibility (`.visible`, `.hidden`, `.veryHidden`) and `activeTab` workbook XML are fully tested
- Sheet protection (`SheetProtection`) including `CoreUtils.configureSheetPassword`, legacy/modern password hashes, and granular permission XML is fully tested
- CLI-generated workbooks from **XLKitTestRunner** are validated with CoreXLSX (see [Test-Workflows/README.md](../Test-Workflows/README.md))

## Test File Structure

The test suite is organized into separate test files by functionality for better maintainability and clarity:

```
Tests/XLKitTests/
├── XLKitTestBase.swift          # Shared XLKitTestSupport helpers (Swift Testing)
├── CoreTests.swift               # Workbook and sheet operations (5 tests)
├── CellValueTests.swift          # Cell values and data types (6 tests)
├── CoordinateTests.swift          # Coordinates and ranges (2 tests)
├── FormattingTests.swift         # Cell formatting (8 tests)
├── NumberFormatTests.swift       # Number formatting (5 tests)
├── TextWrappingTests.swift       # Text wrapping (2 tests)
├── BorderTests.swift             # Border functionality (3 tests)
├── MergeTests.swift              # Cell merging (4 tests)
├── CSVTests.swift                # CSV/TSV operations (12 tests)
├── FileOperationTests.swift      # File operations (2 tests)
├── ImageTests.swift              # Image management (2 tests)
├── ColumnOrderingTests.swift     # Column ordering (2 tests)
├── SheetUtilityTests.swift       # Sheet utilities (6 tests)
├── SheetStateTests.swift         # Sheet visibility state (7 tests)
└── SheetProtectionTests.swift    # Sheet protection (14 tests)
```

### Shared Test Helpers (`XLKitTestSupport` in `XLKitTestBase.swift`)

All test suites are `@Suite` + `@MainActor` structs that call **`XLKitTestSupport`** static helpers:

- **Date Helpers**: `makeUTCDate()` with comprehensive error handling and deterministic fallback, `fixedTestDate`, `epochDate` for deterministic date testing
- **File Helpers**: 
  - `makeTempWorkbookURL()`: Generates UUID-based unique temporary file URLs to prevent concurrent test conflicts
  - `withSavedTempWorkbookSync()`: Creates a standard test workbook, saves it to disk, passes it to the test closure, and ensures cleanup with error logging
  - `withSavedTempWorkbookAsync()`: Async version with the same behavior and cleanup guarantees
  - Shared helpers (`makeTestWorkbook()`, `cleanupTempFile(at:)`) centralize setup/cleanup logic to reduce duplication
- **Format Helpers**: `makeThinBorderFormat()`, `makeMediumRedBorderFormat()`, `makeThickBlueBorderFormat()` for common border configurations, built on a shared bordered-format helper
- **Constants**: `standardFontSize` for consistent font size testing

**Quality Features:**
- **Deterministic Testing**: Fixed dates and UUID-based temp files ensure consistent, reproducible test results
- **Error Visibility**: Cleanup failures are logged with `Issue.record` instead of being silently ignored
- **Safe Error Handling**: Uses `Issue.record` with fallback dates instead of `fatalError` to prevent test suite crashes
- **File Safety**: Workbooks are saved to disk before being passed to test closures, ensuring file operations work correctly
- **Reduced Duplication**: Shared setup/cleanup and bordered-format builders keep test helpers concise and easier to maintain

This structure improves:
- **Maintainability**: Smaller, focused files are easier to review and modify
- **Discoverability**: Tests are organized by functionality, making it easy to find relevant tests
- **Scalability**: New test files can be added without cluttering existing ones
- **Code Reuse**: Shared helpers reduce duplication across test files
- **Test Quality**: Deterministic behavior, proper error handling, and comprehensive cleanup ensure reliable test execution

## Core Workbook & Sheet Tests

**File**: `CoreTests.swift` (5 tests)

- `testCreateWorkbook()`: Basic workbook creation and initialization
- `testAddSheet()`: Sheet addition with proper ID assignment
- `testGetSheetByName()`: Sheet retrieval by name functionality with property validation
- `testRemoveSheet()`: Sheet removal and count tracking
- `testClearSheet()`: Sheet clearing operations (cells, merges, formatting)

## Cell Operations & Data Types

**File**: `CellValueTests.swift` (6 tests)

- `testSetAndGetCell()`: Setting/getting all cell value types (string, number, integer, boolean, date, formula)
- `testSetCellByRowColumn()`: Cell operations using row/column coordinates
- `testSetRange()`: Range-based cell operations
- `testGetUsedCells()`: Used cell detection
- `testCellValueStringValue()`: String representation for all cell values (including date string validation)
- `testCellValueType()`: Type identification for all cell values

## Coordinate & Range Tests

**File**: `CoordinateTests.swift` (2 tests)

- `testCellCoordinate()`: Excel coordinate creation, parsing (including lowercase addresses e.g. "a1", "aa10"), and invalid address handling
- `testCellRange()`: Range creation, notation parsing, and coordinate enumeration

## File Operations

**File**: `FileOperationTests.swift` (2 tests)

- `testWorkbookSave()`: Synchronous workbook saving and file validation using `XLKitTestSupport.withSavedTempWorkbookSync()` (workbook is pre-saved before test execution)
- `testWorkbookSaveAsync()`: Async workbook saving and file validation using `XLKitTestSupport.withSavedTempWorkbookAsync()` (workbook is pre-saved before test execution)

Both tests use **`XLKitTestSupport`** helpers which ensure:
- Workbooks are saved to disk before being passed to test closures
- Automatic cleanup with error logging if cleanup fails
- UUID-based temporary file names to prevent conflicts

## Image & Aspect Ratio Tests

**File**: `ImageTests.swift` (2 tests)

- `testWorkbookImageManagement()`: Workbook-level image operations and management
- `testSheetImageManagement()`: Sheet-level image operations and management

### Aspect Ratio Preservation Testing

**`ImageTests.swift`** covers workbook/sheet image registration in the unit suite. **`XLKitTestRunner`** `embed` validates embedding for all 17 professional video and cinema aspect ratios (see **`Test-Data/Embed-Test/`**):

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

**File**: `FormattingTests.swift` (8 tests)

- `testCellFormatting()`: Basic cell formatting with predefined styles
- `testFontColorFormatting()`: Font colour formatting with proper XML generation and theme colour support
- `testSetCellWithFormatting()`: Cell formatting with custom properties and persistence
- `testHorizontalAlignment()`: All five horizontal alignment options (left, center, right, justify, distributed)
- `testVerticalAlignment()`: All five vertical alignment options (top, center, bottom, justify, distributed)
- `testCombinedAlignment()`: Combined horizontal and vertical alignment scenarios
- `testAlignmentWithOtherFormatting()`: Alignment with other formatting (font, background, etc.)
- `testAlignmentEnumValues()`: Enum value correctness for all alignment options

## Border & Merge Tests

**File**: `BorderTests.swift` (3 tests)

- `testBordersActuallyWork()`: Border functionality with different styles and colors using helper methods
- `testDifferentBorderStyles()`: All border styles (thin, medium, thick) with colors
- `testBorderWithOtherFormatting()`: Borders combined with other formatting options

**File**: `MergeTests.swift` (4 tests)

- `testMergeCells()`: Basic cell merging functionality
- `testMergedCellsActuallyWork()`: Merged cells functionality with complex scenarios
- `testComplexMerging()`: Various merge scenarios (horizontal, vertical, large merges)
- `testComplexBorderAndMergeCombination()`: Combined borders and merges with formatting

## Number Format Tests

**File**: `NumberFormatTests.swift` (5 tests)

- `testNumberFormatCurrency()`: Currency formatting with predefined and custom formats
- `testNumberFormatPercentage()`: Percentage formatting
- `testNumberFormatCustom()`: Custom number format patterns
- `testNumberFormatInFormatToKey()`: Number format inclusion in format keys (tests internal format key generation)
- `testNumberFormatExcelGeneration()`: Excel file generation with number formats

## Text Wrapping Tests

**File**: `TextWrappingTests.swift` (2 tests)

- `testTextWrapping()`: Text wrapping functionality with proper Excel XML generation
- `testTextWrappingInFormatToKey()`: Text wrapping inclusion in format key generation (tests internal format key generation)

### Text Alignment & Wrapping Coverage
- All horizontal and vertical alignment options are tested for correct API behavior and Excel-compliant output.
- Text wrapping functionality is tested with proper Excel XML generation and format key inclusion.
- Combined and formatting-with-alignment scenarios are validated.

## CSV/TSV Import/Export

**File**: `CSVTests.swift` (12 tests)

- `testWorkbookFromCSV()`: Workbook creation from CSV data
- `testWorkbookFromTSV()`: Workbook creation from TSV data
- `testWorkbookImportCSV()`: CSV import into existing workbooks
- `testWorkbookImportTSV()`: TSV import into existing workbooks
- `testWorkbookExportSheetToCSV()`: Export sheet to CSV format
- `testWorkbookExportSheetToTSV()`: Export sheet to TSV format
- `testSheetExportToCSV()`: Sheet-level CSV export
- `testSheetExportToTSV()`: Sheet-level TSV export
- `testCSVWithQuotedFields()`: CSV with quoted fields containing commas (validates swift-textfile integration)
- `testCSVWithEscapedQuotes()`: CSV with escaped quotes (double quotes) - validates proper parsing
- `testCSVExportWithSpecialCharacters()`: CSV export/import round-trip with special characters
- `testCSVWithEmptyFields()`: CSV with empty fields at various positions

## Column & Row Sizing

Column width and row height APIs are tested in **`SheetUtilityTests.swift`** (`testColumnAndRowSizing`). Image-driven cell sizing and aspect-ratio preservation are exercised by **`XLKitTestRunner`** `embed` / `comprehensive` and **`Test-Data/Embed-Test/`**, not by additional unit tests in this target.

## Column Ordering Tests

**File**: `ColumnOrderingTests.swift` (2 tests)

- `testColumnOrderingBeyondZ()`: Excel column order validation for sheets with more than 26 columns (A-Z, AA, AB, etc.)
- `testColumnOrderingWithGaps()`: Column ordering validation with gaps in column sequence

## Sheet Utility Tests

**File**: `SheetUtilityTests.swift` (6 tests)

- `testColumnAndRowSizing()`: Column width and row height management
- `testFluentAPI()`: Fluent API with method chaining
- `testSheetConvenienceMethods()`: Sheet convenience methods for different data types
- `testSheetRowAndColumnMethods()`: Row and column methods for bulk operations
- `testSheetUtilityProperties()`: Sheet utility properties and cell counting
- `testSheetConvenienceInitializer()`: Sheet convenience initializer with data

## Sheet State Tests

**File**: `SheetStateTests.swift` (7 tests)

- `testDefaultStateIsVisible()`: New sheets default to `.visible`
- `testVisibleSheetOmitsStateAttribute()`: Visible sheets emit no `state` attribute so existing files stay byte-identical
- `testHiddenSheetEmitsStateAttribute()`: `.hidden` sheets emit `state="hidden"`
- `testVeryHiddenSheetEmitsStateAttribute()`: `.veryHidden` sheets emit `state="veryHidden"`
- `testActiveTabOmittedWhenFirstSheetVisible()`: No `activeTab` attribute when the first sheet is visible
- `testActiveTabPointsAtFirstVisibleSheet()`: `activeTab` points at the first visible sheet when earlier sheets are hidden
- `testSavingWorkbookWithHiddenSheetSucceeds()`: End-to-end save with a hidden tech sheet writes a valid file

## Sheet Protection Tests

**File**: `SheetProtectionTests.swift` (14 tests)

- `testDefaultProtectionIsNil()`: New sheets have no protection by default
- `testDefaultProtectionStructEnablesSheet()`: `SheetProtection()` defaults to `sheet: true`
- `testLegacyPasswordHashFor1234()`: Legacy hash for plaintext `"1234"` is `CC3D` (Excel-compatible algorithm, not the incorrect OOXML formula)
- `testModernPasswordHashFor1234ComprehensiveDemoPasswordSheet()`: SHA-512 `saltValue` / `hashValue` / `spinCount` for comprehensive-demo **Protected (Password)** sheet salt
- `testModernPasswordHashFor1234ComprehensiveDemoModernSheet()`: SHA-512 vector for comprehensive-demo **Protected (Modern Hash)** sheet salt
- `testConfigureSheetPasswordAppliesLegacyAndModern()`: `CoreUtils.configureSheetPassword` fills legacy `password` and modern fields
- `testConfigureSheetPasswordRejectsEmptyPassword()`: Empty plaintext throws `XLKitError.securityError`
- `testMinimalProtectionXML()`: Default-constructed protection emits `<sheetProtection sheet="1"/>`
- `testProtectionWithLegacyPassword()`: 16-bit `password` attribute is emitted alongside granular flags
- `testProtectionWithModernHash()`: `algorithmName` / `hashValue` / `saltValue` / `spinCount` are emitted for SHA-512 protection
- `testProtectionWithGranularPermissions()`: Boolean flags such as `formatCells`, `insertRows`, `sort` render as `1`/`0`
- `testProtectionAllAttributesEmitted()`: All 21 attribute names appear in the XML when every flag is set; guards against typos in field names
- `testProtectionOmittedWhenNotConfigured()`: Sheets without `protection` set produce no `<sheetProtection>` element
- `testSavingWorkbookWithProtectedSheetSucceeds()`: End-to-end save with a protected sheet writes a valid file

**Demo salt fixtures:** `ComprehensiveDemoProtectionFixtures` in this file must stay aligned with **`Sources/XLKitTestRunner/ComprehensiveDemoProtection.swift`** (used by `comprehensive` CLI and **`Comprehensive-Demo.xlsx`**). See **`Test-Workflows/README.md`** for unprotect password **`1234`**.

## XLKitTestRunner

A modular command-line tool for generating Excel files for testing and demonstration purposes.

### Available Test Types
- `no-embeds` / `no-images`: Generate Excel from CSV without images
- `embed` / `with-embeds` / `with-images`: Generate Excel with embedded images from CSV data
- `comprehensive` / `demo`: 11-sheet API demonstration (formatting, images, CSV patterns, column order beyond Z, sheet visibility, sheet protection; demo password **`1234`** on protected sheets)
- `sheet-password` / `password-hash` / `sheet-password-hash`: Print legacy + SHA-512 `SheetProtection` fields for a plaintext password (optional `--demo-salts` with **`1234`** to match **`Comprehensive-Demo.xlsx`**)
- `security-demo` / `security`: Demonstrate file path security restrictions (console only)
- `ios-test` / `ios`: Test iOS file system compatibility and platform-specific features
- `number-formats` / `formats`: Test number formatting (currency, percentage, custom formats)
- `help` / `-h` / `--help`: Show available commands

### Test Features
- Security Integration: All tests include security logging and validation
- CoreXLSX Validation: Generated files are validated for Excel compliance
- Aspect Ratio Testing: Image embedding tests all 17 professional aspect ratios
- Performance Testing: Large dataset handling and memory optimisation
- Error Handling: Comprehensive error testing and edge case coverage

### Output Structure
```
Test-Workflows/
├── Embed-Test.xlsx          # From no-embeds test
├── Embed-Test-Embed.xlsx    # From embed test (with images)
├── Comprehensive-Demo.xlsx  # From comprehensive test
├── Number-Format-Test.xlsx  # From number-formats test
└── [Your-Test].xlsx         # From custom tests

Root directory:
├── iOS-Example.xlsx         # From ios-test (iOS compatibility)
└── [Other-Test].xlsx        # From other platform-specific tests
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

You can run tests by suite name or by test name pattern:

```bash
# Run tests from a specific suite
swift test --filter CoreTests
swift test --filter CSVTests
swift test --filter BorderTests

# Image-related tests
swift test --filter "test.*Image.*|test.*Aspect.*"

# CSV/TSV tests
swift test --filter "test.*CSV.*|test.*TSV.*"

# Core functionality tests
swift test --filter "testCreateWorkbook|testAddSheet|testSetAndGetCell"

# Formatting tests
swift test --filter FormattingTests

# Border and merge tests
swift test --filter "BorderTests|MergeTests"

# Sheet visibility and protection
swift test --filter SheetStateTests
swift test --filter SheetProtectionTests
swift test --filter "testLegacyPasswordHashFor1234|testConfigureSheetPassword"
```

### Running XLKitTestRunner
```bash
# Run specific test types
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner sheet-password 1234 --demo-salts
swift run XLKitTestRunner help
```

### Validation
- All tests generate valid Excel files
- CoreXLSX validation for Excel compliance
- Aspect ratio tests verify pixel-perfect preservation
- File operations tests validate complete workflows
- Security features are validated in all operations

## Coverage & Quality Assurance

| Category           | Test File              | Test Count | Coverage         |
|--------------------|------------------------|------------|-----------------|
| Core Workbook      | CoreTests.swift        | 5          | 100%            |
| Cell Values        | CellValueTests.swift   | 6          | 100%            |
| Coordinates/Ranges | CoordinateTests.swift  | 2          | 100%            |
| File Operations    | FileOperationTests.swift | 2        | 100%            |
| Image Support      | ImageTests.swift        | 2          | 100%            |
| Cell Formatting    | FormattingTests.swift   | 8          | 100%            |
| Border Tests       | BorderTests.swift       | 3          | 100%            |
| Merge Tests        | MergeTests.swift        | 4          | 100%            |
| Number Formats     | NumberFormatTests.swift | 5          | 100%            |
| Text Wrapping      | TextWrappingTests.swift | 2          | 100%            |
| CSV/TSV            | CSVTests.swift          | 12         | 100%            |
| Column Ordering    | ColumnOrderingTests.swift | 2        | 100%            |
| Sheet Utilities    | SheetUtilityTests.swift | 6          | 100%            |
| Sheet State        | SheetStateTests.swift   | 7          | 100%            |
| Sheet Protection   | SheetProtectionTests.swift | 14      | 100%            |
| **Total**          | **15 test files**      | **80**     | **100%**        |

### Quality Standards
- All generated files pass CoreXLSX validation
- Zero tolerance for aspect ratio distortion
- Efficient memory usage and file generation
- Comprehensive error and edge case testing
- 100% of public APIs tested
- Deterministic test execution: fixed dates and UUID-based temp files ensure reproducible results
- Proper error handling: no force unwraps, comprehensive guard statements, detailed error messages
- Test isolation: unique temporary files prevent concurrent test conflicts
- Cleanup reliability: file cleanup failures are logged instead of silently ignored
- All text alignment options (horizontal, vertical, combined) are fully tested
- All text wrapping functionality is fully tested
- All border and merge functionality is fully tested
- All number formatting options are fully tested
- Column ordering for sheets with more than 26 columns is fully tested
- Sheet visibility and sheet protection (including password hash helpers) are fully tested
- Automated CI on macOS and iOS (`build.yml`: unit tests, strict-concurrency job, TestRunner smoke; CLI workflows run XLKitTestRunner)
- Security features integrated throughout all tests
- XLKitTestRunner provides comprehensive demonstration capabilities

### Security Integration
- Rate limiting prevents test abuse
- Security logging captures all operations
- Input validation ensures data integrity
- File quarantine protects against malicious content
- Checksum verification ensures file authenticity

## More documentation

- **`Documentation/Manual/10-Testing-Test-Runner-CI-and-Code-Style.md`** — Unit tests, CI workflows, code style
- **`Documentation/Manual/03-Core-Model-Workbook-Sheet-and-Cells.md`** — `SheetState`, `SheetProtection`
- **`Documentation/Manual/09-Errors-CoreUtils-and-iOS.md`** — `configureSheetPassword`, password hash helpers
- **`Test-Workflows/README.md`** — Generated `.xlsx` outputs and comprehensive demo passwords
- **`Sources/XLKitTestRunner/README.md`** — Full CLI reference
- **`AGENT.MD`** — Contributor and AI-agent expectations

This suite ensures XLKit delivers reliable, high-quality Excel file generation with perfect image embedding and aspect ratio preservation, backed by comprehensive security features and validation. 