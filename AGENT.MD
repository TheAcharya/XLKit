# XLKit - AI Agent Development Guide

## Table of Contents

- [Project Overview](#project-overview)
- [Architecture](#architecture)
  - [Module Structure](#module-structure)
  - [Dependencies](#dependencies)
  - [Test Runner](#test-runner)
- [Development Standards](#development-standards)
  - [Code Quality](#code-quality)
  - [Security Requirements](#security-requirements)
  - [Security Features](#security-features)
  - [Testing Strategy](#testing-strategy)
  - [Documentation](#documentation)
- [API Reference](#api-reference)
  - [Core Types](#core-types)
  - [Sheet Operations](#sheet-operations)
  - [Workbook Operations](#workbook-operations)
  - [CSV/TSV Operations](#csvtsv-operations)
  - [Image Operations](#image-operations)
- [Recent Improvements](#recent-improvements)
  - [iOS Compatibility Fix](#ios-compatibility-fix-2025-07-14)
  - [Perfect Aspect Ratio Preservation](#perfect-aspect-ratio-preservation-2025-07-07)
  - [Scaling API Investigation and Fixes](#scaling-api-investigation-and-fixes-2025-07-12)
- [Implementation Details](#implementation-details)
  - [XLSX Generation](#xlsx-generation)
  - [Image Embedding Scaling API](#image-embedding-scaling-api)
  - [Image Embedding](#image-embedding)
  - [Error Handling](#error-handling)
- [Performance & Optimization](#performance--optimization)
- [Maintenance & Updates](#maintenance--updates)

---

## Project Overview

XLKit is a modern Swift library for creating and manipulating Excel (.xlsx) files on macOS and iOS. Built with Swift 6.0, targeting macOS 12+ and iOS 15+, using modular Swift Package Manager architecture. iOS support is available and tested in CI/CD, with platform-specific code handling for iOS compatibility.

## Architecture

### Module Structure

The library is organized into five SPM modules:

- **XLKitCore**: Core types, data structures, and utilities
- **XLKitFormatters**: CSV/TSV import/export functionality  
- **XLKitImages**: Image processing and embedding utilities
- **XLKitXLSX**: XLSX file generation engine
- **XLKit**: Main API that re-exports all submodules

### Directory and File Structure

```
XLKit/
├── AGENT.MD                     # AI agent development guide
├── .cursorrules                 # Cursor rules for AI agents
├── CHANGELOG.md                 # Version history and changes
├── LICENSE                      # MIT license
├── Package.swift                # Swift Package Manager configuration
├── Package.resolved             # Locked dependency versions
├── README.md                    # Main documentation
├── SECURITY.md                  # Security policy
├── .gitignore                   # Git ignore patterns
├── .swift-format                # Swift formatting configuration
├── Assets/                      # Project assets
│   └── XLKit_Icon.png          # Project icon
├── Sources/                     # Source code modules
│   ├── XLKit/                  # Main API module
│   │   ├── XLKit.swift         # Main API exports
│   │   ├── Sheet+API.swift     # Sheet operations API
│   │   └── Workbook+API.swift  # Workbook operations API
│   ├── XLKitCore/              # Core types and utilities
│   │   ├── CoreTypes.swift     # Core data structures (1215 lines)
│   │   └── SecurityManager.swift # Security features (282 lines)
│   ├── XLKitFormatters/        # CSV/TSV functionality
│   │   └── CSVUtils.swift      # CSV import/export utilities (199 lines)
│   ├── XLKitImages/            # Image processing
│   │   ├── ImageUtils.swift    # Image utilities (155 lines)
│   │   └── ImageSizingUtils.swift # Image sizing logic (191 lines)
│   ├── XLKitXLSX/              # XLSX generation engine
│   │   └── XLSXEngine.swift    # XLSX file generation (897 lines)
│   └── XLKitTestRunner/        # Test runner executable
│       ├── main.swift          # Command-line interface (91 lines)
│       ├── ExcelGenerators.swift # Excel generation tests (590 lines)
│       ├── ImageEmbedGenerators.swift # Image embedding tests (228 lines)
│       ├── README.md           # Test runner documentation (199 lines)
│       └── Templates/          # Test templates
│           └── TestGeneratorTemplate.swift # Template for new tests (224 lines)
├── Tests/                      # Unit tests
│   ├── README.md               # Test documentation (196 lines)
│   └── XLKitTests/             # Test suite
│       └── XLKitTests.swift    # All unit tests (1430 lines, 55 tests)
├── Test-Data/                  # Test data files
│   ├── README.md               # Test data documentation (44 lines)
│   └── Embed-Test/             # Image embedding test data
│       ├── Embed-Test.csv      # CSV test data (5 lines)
│       ├── Embed-Test_00-00-08-06.png # Test image 1 (681KB)
│       ├── Embed-Test_00-00-22-07.png # Test image 2 (825KB)
│       ├── Embed-Test_00-00-50-08.png # Test image 3 (779KB)
│       └── Embed-Test_00-01-09-10.png # Test image 4 (703KB)
├── Test-Workflows/             # Generated Excel files
│   └── README.md               # Output documentation
└── .github/                    # GitHub configuration
    ├── FUNDING.yml             # Funding configuration
    └── workflows/              # CI/CD workflows
        ├── build.yml           # Main build and test workflow (85 lines)
        ├── codeql.yml          # Security scanning workflow (116 lines)
        ├── cli-embed.yml       # Image embedding test workflow (37 lines)
        ├── cli-generic.yml     # Generic test workflow (41 lines)
        ├── cli-no-embeds.yml   # No-embeds test workflow (37 lines)
        └── cli-ios.yml         # iOS compatibility test workflow (37 lines)
```

### Dependencies

```
XLKit (main API)
├── XLKitCore (core types & utilities)
├── XLKitFormatters (CSV/TSV import/export)
├── XLKitImages (image processing & embedding)
└── XLKitXLSX (XLSX generation engine)

XLKitFormatters
└── XLKitCore

XLKitImages
└── XLKitCore

XLKitXLSX
├── XLKitCore
├── XLKitFormatters
└── XLKitImages
```

External Dependencies:
- CoreXLSX (0.14.2): Excel file validation and parsing
- ZIPFoundation (0.9.19): Cross-platform ZIP archive creation
- XMLCoder (0.14.0): XML serialization

### Test Runner

XLKitTestRunner is a modular command-line tool for generating Excel files for testing and demonstration purposes.

Structure:
```
Sources/XLKitTestRunner/
├── main.swift                    # Entry point with CLI
├── ExcelGenerators.swift         # Excel generation functions
├── ImageEmbedGenerators.swift    # Image embedding tests
├── Templates/                    # Template files
│   └── TestGeneratorTemplate.swift
└── README.md                     # Documentation
```

Usage:
```bash
# Run specific test types
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner help
```

Available Test Types:
- `no-embeds` / `no-images`: Generate Excel from CSV without images
- `embed` / `with-embeds` / `with-images`: Generate Excel with embedded images from CSV data
- `comprehensive` / `demo`: Comprehensive API demonstration with all features
- `security-demo` / `security`: Demonstrate file path security restrictions
- `ios-test` / `ios`: Test iOS file system compatibility and platform-specific features
- `number-formats` / `formats`: Test number formatting (currency, percentage, custom formats)
- `help` / `-h` / `--help`: Show available commands

Test Features:
- Security Integration: All tests include security logging and validation
- CoreXLSX Validation: Generated files are validated for Excel compliance
- Aspect Ratio Testing: Image embedding tests all 17 professional aspect ratios
- Performance Testing: Large dataset handling and memory optimization
- Error Handling: Comprehensive error testing and edge case coverage
- Platform Testing: iOS compatibility validation and sandbox restrictions testing

Output Structure:
```
Test-Workflows/
├── Embed-Test.xlsx          # From no-embeds test
├── Embed-Test-Embed.xlsx    # From embed test (with images)
├── Comprehensive-Demo.xlsx  # From comprehensive test
├── Number-Format-Test.xlsx  # From number-formats test
└── [Your-Test].xlsx         # From custom tests

Root Directory:
├── iOS-Example.xlsx         # From ios-test (iOS compatibility)
└── [Other-Test].xlsx        # From other platform-specific tests
```

Security Features in Tests:
- Rate Limiting: Prevents test abuse and resource exhaustion
- Security Logging: All test operations are logged for audit trails
- Input Validation: All test inputs are validated for security
- File Quarantine: Suspicious test files are automatically quarantined
- Checksum Verification: Optional file integrity verification (disabled by default)

## Development Standards

### Code Quality

Swift 6.0 Compliance:
- Use `@preconcurrency` imports for modules with Sendable types
- Note: Workbook and Sheet classes are not Sendable due to mutable state requirements
- Use modern Swift idioms and features
- Avoid force-unwraps and force-casts in public APIs

Cross-Platform Compatibility:
- Use platform-specific conditionals (`#if os(macOS)`, `#if os(iOS)`) for platform-specific APIs
- Avoid iOS-unavailable APIs like `FileManager.default.homeDirectoryForCurrentUser`
- Test builds on both macOS and iOS platforms
- Ensure all file operations work on both platforms

Code Style:
- Use 4-space indentation (no tabs)
- Use trailing commas for better git diffs
- Group and reorder imports alphabetically
- Use MARK comments for code organization

File Structure:
```swift
//
//  Filename.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
@preconcurrency import XLKitCore

// MARK: - Section Name
// Implementation
```

Error Handling:
- Use specific XLKitError types
- Provide meaningful error messages
- Use guard statements for early returns
- Handle errors gracefully in public APIs

Type Safety:
- Use strong typing throughout
- Prefer enums over strings for constants
- Use structs for value types, classes for reference types
- Implement Equatable, Hashable where appropriate

### Security Requirements

All code must be free from vulnerability injection, supply chain poisoning, and hidden malicious logic:

- No dynamic code execution, system shell commands, or unsafe reflection
- All dependencies must be reputable, open-source, and version-pinned in Package.resolved
- No network, HTTP, or remote code execution except for documented, safe APIs
- All file operations must be local and safe, with no writing to unexpected or system locations
- No hardcoded secrets, tokens, or suspicious data
- All code must be readable, idiomatic Swift, and match the documented architecture
- Every code change must pass static analysis and security review

### Security Features

XLKit includes comprehensive security features to protect against vulnerabilities, supply chain attacks, and malicious code injection:

#### SecurityManager

The `SecurityManager` provides centralized security controls and monitoring:

```swift
@MainActor
public struct SecurityManager {
    // Configuration
    public static var enableChecksumStorage = false  // Disabled by default
    
    // Rate limiting, logging, quarantine, and checksum features
}
```

#### Rate Limiting

Prevents abuse and resource exhaustion:
- Default: 100 operations per minute
- Configurable: Time window and operation limits
- Automatic: Integrated into all file operations
- Logging: Rate limit violations are logged

```swift
// Check rate limit before operations
try SecurityManager.checkRateLimit()
```

#### Security Logging

Comprehensive audit trail of all security-relevant operations:
- Console logging: Real-time security events
- File logging: Persistent audit trail in temporary directory
- Structured data: Timestamp, operation, details, user agent
- Operations logged: File generation, checksums, quarantines, rate limits

```swift
// Log security operations
SecurityManager.logSecurityOperation("xlsx_generation_started", details: [
    "target_path": filePath,
    "workbook_sheets": sheetCount,
    "workbook_images": imageCount
])
```

#### File Quarantine

Isolates suspicious files to prevent execution:
- Pattern detection: Checks for malicious code patterns
- Size validation: Prevents oversized files
- Format validation: Validates image formats
- Automatic isolation: Moves suspicious files to quarantine directory

```swift
// Check if file should be quarantined
if SecurityManager.shouldQuarantineFile(imageData, format: .png) {
    try SecurityManager.quarantineSuspiciousFile(fileURL, reason: "Suspicious content detected")
}
```

#### File Checksums

Cryptographic integrity verification (optional):
- SHA-256 hashes: Secure file integrity verification
- Configurable: Can be enabled/disabled via `enableChecksumStorage`
- Tamper detection: Identifies unauthorized file modifications
- Supply chain protection: Ensures file authenticity

```swift
// Store file checksum (when enabled)
SecurityManager.storeFileChecksum(checksum, for: fileURL)

// Verify file integrity (when enabled)
let isValid = SecurityManager.verifyFileChecksum(fileURL)
```

#### Input Validation

Comprehensive validation of all user inputs:
- File paths: Validates and sanitizes file paths
- Image data: Validates image formats and sizes
- CSV data: Validates CSV structure and content
- Coordinates: Validates Excel coordinate formats

#### Error Handling

Secure error handling that doesn't expose sensitive information:
- Generic messages: User-friendly error messages
- Detailed logging: Internal error details logged for debugging
- Graceful degradation: Continues operation when possible
- Security events: Security-related errors trigger additional logging

#### Security Integration

Security features are integrated throughout the codebase:
- XLSXEngine: Rate limiting, logging, checksums for file generation
- ImageUtils: Quarantine, validation for image processing
- XLKit API: Input validation, security logging for all operations
- Test Runner: Security validation for all test operations

#### Configuration

Security features can be configured as needed:
```swift
// Enable checksum storage (disabled by default)
SecurityManager.enableChecksumStorage = true

// Security logging is always active
// Rate limiting is always active
// Input validation is always active
// File quarantine is always active
```

#### Security Log Output

Example security log entries:
```
[SECURITY] 2025-07-08 8:25:48 PM +0000: xlsx_generation_started - ["target_path": "...", "workbook_sheets": 1, "workbook_images": 0]
[SECURITY] 2025-07-08 8:25:48 PM +0000: xlsx_generation_completed - ["checksum": "...", "file_size": 19676, "target_path": "..."]
[SECURITY] 2025-07-08 8:25:48 PM +0000: checksum_stored - ["checksum": "...", "timestamp": 1752006230.891748, "file_path": "..."]
```

### Testing Strategy

Test Coverage Requirements:
- Test all public APIs
- Test error conditions and edge cases
- Test CSV/TSV import/export functionality
- Test image format detection and embedding with all 17 aspect ratios
- Test XLSX file generation and saving
- Test coordinate and range operations
- Test font colour formatting and XML generation
- Test all text alignment options (horizontal, vertical, combined)
- Test text wrapping functionality with proper Excel XML generation
- Test border functionality with different styles and colors
- Test merged cells with complex scenarios
- Test number formatting (currency, percentage, custom formats)
- Test column ordering for sheets with more than 26 columns (A-Z, AA, AB, etc.)
- Test platform compatibility and iOS-specific features

Test Patterns:
```swift
func testFeatureName() {
    // Arrange
    let workbook = Workbook()
    
    // Act
    let result = workbook.someOperation()
    
    // Assert
    XCTAssertEqual(result, expectedValue)
}

func testFeatureNameWithInvalidInput() {
    // Act & Assert
    XCTAssertThrowsError(try someOperation(invalidInput)) { error in
        XCTAssertEqual(error as? XLKitError, .expectedError)
    }
}
```

Performance Testing:
- Test with large datasets
- Test memory usage patterns
- Test async operations
- Test concurrent access

Text Alignment Testing:
- Test all 5 horizontal alignment options (left, center, right, justify, distributed)
- Test all 5 vertical alignment options (top, center, bottom, justify, distributed)
- Test combined horizontal and vertical alignment scenarios
- Test alignment with other formatting options (font, background, etc.)
- Test enum value correctness for all alignment options
- Test format key generation includes alignment information
- Test Excel-compliant XML generation for all alignment options

Border and Merge Testing:
- Test all border styles (thin, medium, thick) with different colors
- Test border combinations with other formatting options
- Test merged cells with complex scenarios (horizontal, vertical, large merges)
- Test border and merge combinations with formatting
- Test Excel-compliant XML generation for borders and merges
- Test format key generation includes border information
- Test text wrapping functionality with proper Excel XML generation
- Test text wrapping inclusion in format key generation
- Test column ordering with proper Excel column sequence (A, B, ..., Z, AA, AB, ...)

### Documentation

Code Documentation:
- All public APIs must have doc comments
- Include parameter descriptions
- Provide usage examples
- Document error conditions

README Documentation:
- Keep README.md comprehensive and up-to-date
- Include quick start examples
- Document all major features
- Provide API reference sections

Test Documentation:
- All test documentation maintained in Tests/README.md
- Must accurately reflect all current tests and coverage
- Update when tests are added, removed, or reorganized

## API Reference

### Core Types

Workbook Class:
```swift
public final class Workbook {
    private var sheets: [Sheet] = []
    private let nextSheetId: Int
    private var images: [ExcelImage] = []
}
```
- Manages multiple sheets and workbook-level images
- Auto-increments sheet IDs
- Not Sendable due to mutable state requirements

Sheet Class:
```swift
public final class Sheet: Equatable {
    public let name: String
    public let id: Int
    private var cells: [String: CellValue] = [:]
    private var mergedRanges: [CellRange] = []
    private var columnWidths: [Int: Double] = [:]
    private var rowHeights: [Int: Double] = [:]
    private var images: [String: ExcelImage] = [:]
    private var cellFormats: [String: CellFormat] = [:]
}
```
- Represents a worksheet with cells, formatting, and images
- Supports cell operations, range operations, and image embedding
- Not Sendable due to mutable state requirements

CellValue Enum:
```swift
public enum CellValue: Equatable {
    case string(String)
    case number(Double)
    case integer(Int)
    case boolean(Bool)
    case date(Date)
    case formula(String)
    case empty
}
```

CellCoordinate Struct:
```swift
public struct CellCoordinate: Hashable {
    public let row: Int
    public let column: Int
    public var excelAddress: String // e.g., "A1", "B2"
}
```

CellRange Struct:
```swift
public struct CellRange: Hashable {
    public let start: CellCoordinate
    public let end: CellCoordinate
    public var coordinates: [CellCoordinate]
    public var excelRange: String // e.g., "A1:B3"
}
```

Cell Struct:
```swift
public struct Cell {
    public let value: CellValue
    public let format: CellFormat?
    
    // Factory methods for each cell type
    public static func string(_ value: String, format: CellFormat? = nil) -> Cell
    public static func number(_ value: Double, format: CellFormat? = nil) -> Cell
    public static func integer(_ value: Int, format: CellFormat? = nil) -> Cell
    public static func boolean(_ value: Bool, format: CellFormat? = nil) -> Cell
    public static func date(_ value: Date, format: CellFormat? = nil) -> Cell
    public static func formula(_ value: String, format: CellFormat? = nil) -> Cell
}
```

CellFormat Struct:
```swift
public struct CellFormat {
    public var fontName: String?
    public var fontSize: Double?
    public var fontWeight: FontWeight?
    public var fontStyle: FontStyle?
    public var fontColor: String?
    public var backgroundColor: String?
    public var horizontalAlignment: HorizontalAlignment?
    public var verticalAlignment: VerticalAlignment?
    public var textDecoration: TextDecoration?
    public var textWrapping: Bool?
    public var textRotation: Int?
    public var numberFormat: NumberFormat?
    public var customNumberFormat: String?
    public var borderTop: BorderStyle?
    public var borderBottom: BorderStyle?
    public var borderLeft: BorderStyle?
    public var borderRight: BorderStyle?
    public var borderColor: String?
    
    // Predefined format factories
    public static func header(fontSize: Double = 14.0, backgroundColor: String = "#E0E0E0") -> CellFormat
    public static func currency(format: NumberFormat = .currencyWithDecimals, color: String? = nil) -> CellFormat
    public static func percentage(format: NumberFormat = .percentageWithDecimals) -> CellFormat
    public static func date(format: NumberFormat = .date) -> CellFormat
    public static func bordered(style: BorderStyle = .thin, color: String? = nil) -> CellFormat
    public static func coloredText(color: String) -> CellFormat
}
```

Border and Merge Support:
XLKit provides comprehensive border and merge functionality:
- Border Support: Thin, medium, and thick border styles with custom colors
- Merge Support: Cell merging with complex range scenarios
- Combined Formatting: Borders and merges work with all other formatting options
- Excel Compliance: Proper XML generation for borders and merges

Text Alignment Support:
XLKit provides comprehensive text alignment support with all 6 alignment options available in Excel:
- Horizontal Alignment: left, center, right, justify, distributed
- Vertical Alignment: top, center, bottom, justify, distributed
- Combined alignment scenarios for precise cell formatting
- Excel-compliant XML generation for all alignment options

ExcelImage Struct:
```swift
public struct ExcelImage {
    public let id: String
    public let data: Data
    public let format: ImageFormat
    public let originalSize: CGSize
    public let displaySize: CGSize?
}
```

ImageFormat Enum:
```swift
public enum ImageFormat: String, CaseIterable {
    case gif = "gif"
    case png = "png"
    case jpeg = "jpeg"
    case jpg = "jpg"
    case bmp = "bmp"
    case tiff = "tiff"
    
    public var mimeType: String
    public var excelContentType: String
}
```

XLKitError Enum:
```swift
public enum XLKitError: Error, LocalizedError {
    case invalidCoordinate(String)
    case invalidRange(String)
    case fileWriteError(String)
    case zipCreationError(String)
    case xmlGenerationError(String)
}
```

### Sheet Operations

Cell Operations:
```swift
// Set cell values
sheet.setCell("A1", string: "Hello World")
sheet.setCell("B1", number: 42.5)
sheet.setCell("C1", integer: 100)
sheet.setCell("D1", boolean: true)
sheet.setCell("E1", date: Date())
sheet.setCell("F1", formula: "=SUM(A1:E1)")

// Get cell values
let value = sheet.getCell("A1")

// Set cell with format
sheet.setCell("A1", cell: Cell.string("Header", format: CellFormat.header()))

// Set cell with font colour
sheet.setCell("A1", cell: Cell.string("Red Text", format: CellFormat.coloredText(color: "#FF0000")))

// Set range of cells
sheet.setRange("A1:A10", value: .string("Batch"))
```

Range Operations:
```swift
// Merge cells
sheet.mergeCells("A1:B2")

// Get merged ranges
let ranges = sheet.getMergedRanges()

// Check if cell is in merged range
let isMerged = sheet.isCellMerged("A1")

// Border formatting
var borderedFormat = CellFormat.bordered()
borderedFormat.borderTop = .thin
borderedFormat.borderBottom = .thin
borderedFormat.borderLeft = .thin
borderedFormat.borderRight = .thin
borderedFormat.borderColor = "#000000"
sheet.setCell("A1", string: "Bordered Cell", format: borderedFormat)

// Complex border and merge combination
var complexFormat = CellFormat.bordered()
complexFormat.fontSize = 11
complexFormat.fontWeight = .bold
complexFormat.horizontalAlignment = .center
complexFormat.verticalAlignment = .center
complexFormat.fontName = "Calibri"
complexFormat.borderTop = .thin
complexFormat.borderBottom = .thin
complexFormat.borderLeft = .thin
complexFormat.borderRight = .thin
complexFormat.borderColor = "#000000"

sheet.setCell("A1", string: "Test1", format: complexFormat)
sheet.setCell("A2", string: "Test2")
sheet.mergeCells("A1:B1")
sheet.mergeCells("A2:B2")
```

Formatting Operations:
```swift
// Set column width
sheet.setColumnWidth(1, width: 15.0)

// Set row height
sheet.setRowHeight(1, height: 20.0)

// Get column width and row height
let width = sheet.getColumnWidth(1)
let height = sheet.getRowHeight(1)
```

Image Operations:
```swift
// Add image with automatic sizing
try sheet.embedImageAutoSized(
    imageData: imageData,
    at: "B2",
    workbook: workbook
)

// Add image from URL
try sheet.embedImageAutoSized(
    imageURL: imageURL,
    at: "B2",
    workbook: workbook
)

// Get all images
let images = sheet.getImages()

// Check if cell has image
let hasImage = sheet.hasImage(at: "B2")
```

Utility Operations:
```swift
// Get used cells
let usedCells = sheet.getUsedCells()

// Clear all data
sheet.clear()
```

### Workbook Operations

Sheet Management:
```swift
// Add sheets
let sheet1 = workbook.addSheet(name: "Sheet1")
let sheet2 = workbook.addSheet(name: "Sheet2")

// Get sheets
let allSheets = workbook.getSheets()
let specificSheet = workbook.getSheet(name: "Sheet1")

// Remove sheets
workbook.removeSheet(name: "Sheet1")
```

Image Management:
```swift
// Add workbook-level images
workbook.addImage(excelImage)

// Get all images
let images = workbook.getImages()
```

### CSV/TSV Operations

Import Operations:
```swift
// Create workbook from CSV
let workbook = Workbook(fromCSV: "Name,Age\nJohn,25\nJane,30", hasHeader: true)

// Import into existing sheet
sheet.importCSV("Name,Age\nJohn,25\nJane,30", hasHeader: true)
```

Export Operations:
```swift
// Export sheet to CSV
let csv = sheet.exportToCSV()

// Export sheet to TSV
let tsv = sheet.exportToTSV()
```

### Image Operations

Image Embedding:
```swift
// Embed image with automatic sizing and perfect aspect ratio
try await sheet.embedImageAutoSized(imageData, at: "B2", of: workbook)

// Embed image with custom scaling
try await sheet.embedImageAutoSized(
    imageData,
    at: "B2",
    of: workbook,
    scale: 0.7
)
```

## Recent Improvements

### Border and Merge Functionality (2025-08-04)
- Comprehensive border and merge functionality implemented and tested
- Border support with thin, medium, and thick styles with custom colors
- Merged cells support with complex range scenarios
- Border and merge combinations with other formatting options
- Excel-compliant XML generation for borders and merges
- Test count increased from 45 to 51 tests with 100% API coverage
- All border and merge functionality fully tested and validated

### Text Wrapping Functionality (2025-09-25)
- Comprehensive text wrapping functionality implemented and tested
- Text wrapping support with proper Excel XML generation
- Text wrapping inclusion in format key generation for proper format caching
- Excel-compliant XML generation for text wrapping with wrapText attribute
- Test count increased from 51 to 53 tests with 100% API coverage
- All text wrapping functionality fully tested and validated

### Column Ordering Fix (2025-10-18)
- Fixed critical column ordering bug for sheets with more than 26 columns (A-Z, AA, AB, etc.)
- Resolved Excel compatibility issue where generated files were rejected or repaired due to invalid column ordering
- Implemented proper numeric column sorting in XLSXEngine.generateWorksheetXML() to ensure Excel-compliant column order
- Fixed lexicographic string sorting that caused "AA1" to appear before "B1" in generated XML
- Added comprehensive test coverage with testColumnOrderingBeyondZ() and testColumnOrderingWithGaps() test functions
- Increased test count from 53 to 55 tests with 100% API coverage including column ordering validation
- Ensured proper Excel column order: A, B, ..., Z, AA, AB, ..., BA, BB, ... in all generated files
- Maintained full backward compatibility with existing APIs and storage model
- Validated fix with CoreXLSX to ensure all generated Excel files open correctly in Excel
- Enhanced XLKitTestRunner with column ordering validation for continuous testing

### iOS Compatibility Fix (2025-07-14)
- Fixed iOS build error: `'homeDirectoryForCurrentUser' is unavailable in iOS`
- Implemented platform-specific conditionals for file system operations
- Updated `allowedDirectories` in CoreTypes.swift to use `#if os(macOS)` for home directory access
- Maintained security features while ensuring cross-platform compatibility
- Verified successful builds on both macOS and iOS platforms
- Added iOS job to GitHub Actions workflow for continuous testing

### Perfect Aspect Ratio Preservation (2025-07-07)
- Implemented pixel-perfect image embedding with zero distortion
- Empirically derived formulas: Column width `pixels / 8.0`, Row height `pixels / 1.33`
- ImageSizingUtils: Centralized sizing logic with consistent formulas
- EMU Coordinate System: Correct Excel internal format (1 pixel = 9525 EMUs)
- Comprehensive testing of all 17 professional video and cinema aspect ratios
- Excel compliance with CoreXLSX validation for all generated files
- Simplified API with automatic sizing and aspect ratio preservation
- Font Colour Support: Added comprehensive font colour formatting with proper XML generation and theme colour support

### Scaling API Investigation and Fixes (2025-07-12)
- Identified and resolved scaling inconsistencies between XLKit test and MarkersExtractor integration
- Fixed XLKit test to use default parameters instead of manual overrides
- Established consistent API usage pattern: let XLKit handle sizing automatically
- Documented scale parameter options and best practices for image embedding
- Verified perfect aspect ratio preservation across all implementations
- Confirmed Excel compliance and CoreXLSX validation for all generated files

## Implementation Details

### XLSX Generation

The XLSX engine generates all required files for Excel compliance:
- docProps (core.xml, app.xml)
- theme (theme1.xml)
- styles (styles.xml)
- sharedStrings (sharedStrings.xml)
- xl/workbook.xml
- xl/worksheets/sheet1.xml
- xl/drawings/drawing1.xml (if images present)
- xl/media/ (image files)
- _rels/ (relationship files)
- [Content_Types].xml

Key Features:
- Uses ZIPFoundation for cross-platform ZIP creation
- Generates Excel-compliant XML with proper escaping
- Validates files using CoreXLSX
- Supports all Excel data types and formatting
- Generates proper font colour XML with theme colour support

### Image Embedding Scaling API

XLKit provides automatic image scaling with perfect aspect ratio preservation. The `embedImageAutoSized` method handles all sizing automatically.

#### Default Parameters
```swift
func embedImageAutoSized(
    _ data: Data,
    at coordinate: String,
    of workbook: Workbook,
    format: ImageFormat? = nil,
    maxCellWidth: CGFloat = 600,    // Default maximum width
    maxCellHeight: CGFloat = 400,   // Default maximum height
    scale: CGFloat = 0.5            // Default 50% scaling
) throws -> Bool
```

#### Scale Control
The `scale` parameter controls image size relative to maximum bounds:
- `scale: 0.3` - 30% (very small images)
- `scale: 0.5` - 50% (default, compact)
- `scale: 0.7` - 70% (medium size)
- `scale: 0.8` - 80% (larger images)
- `scale: 1.0` - 100% (full size, maximum bounds)

#### Automatic Sizing Process
1. XLKit calculates display size within specified bounds
2. Maintains perfect aspect ratio with zero distortion
3. Automatically sets column width using `pixels / 8.0` formula
4. Automatically sets row height using `pixels / 1.33` formula
5. Handles all Excel compliance and EMU calculations

#### Integration Best Practices
- Let XLKit handle all sizing automatically using defaults
- Call `embedImageAutoSized` after column width adjustments
- XLKit will override image column with perfect sizing
- Use `scale` parameter to control image size when needed
- Avoid manual column width calculations for image columns

### Image Embedding

XLKit supports pixel-perfect image embedding with automatic sizing and perfect aspect ratio preservation.

Supported Aspect Ratios (Tested & Validated):
All 17 professional video and cinema aspect ratios with pixel-perfect preservation:
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

Key Features:
- Perfect Aspect Ratio Preservation: Images maintain exact original proportions with zero distortion
- Automatic Cell Sizing: Row height and column width adjust to match image dimensions using empirically derived formulas
- Precise Positioning: Images positioned exactly at cell boundaries with zero offsets
- Excel Compliance: Uses precise Excel formulas for sizing calculations with CoreXLSX validation
- Workbook Registration: Images registered with both sheet and workbook for proper tracking
- Professional Quality: Pixel-perfect exports suitable for professional video editing workflows
- CRITICAL: The pixel-perfect image embedding is the most critical feature and must be preserved at all costs

Implementation:
- Uses ImageSizingUtils for centralized sizing logic with consistent formulas
- Column width: `pixels / 8.0` (empirically derived from manual Excel analysis)
- Row height: `pixels / 1.33` (empirically derived from manual Excel analysis)
- EMU conversion: `pixels * 9525` (Excel's internal format for perfect positioning)
- Drawing XML with correct EMU coordinates and zero offsets
- CoreXLSX validation ensures full Excel compliance for all generated files

### Error Handling

XLKitError Types:
```swift
// Coordinate errors
XLKitError.invalidCoordinate("Z999999") // Invalid Excel address

// Range errors
XLKitError.invalidRange("A1:Z999999") // Invalid range

// File operation errors
XLKitError.fileWriteError("Permission denied")
XLKitError.zipCreationError("ZIP creation failed")
XLKitError.xmlGenerationError("XML generation failed")
```

Error Handling Patterns:
```swift
do {
    try await workbook.save(to: url)
} catch XLKitError.invalidCoordinate(let coord) {
    print("Invalid coordinate: \(coord)")
} catch XLKitError.fileWriteError(let message) {
    print("File write error: \(message)")
} catch {
    print("Unexpected error: \(error)")
}
```

## Performance & Optimization

Memory Management:
- Optimize for large datasets
- Use efficient data structures
- Minimize memory allocations
- Handle cleanup properly

Async Operations:
- Use async/await for file I/O
- Note: Async operations use synchronous implementation since Workbook/Sheet are not Sendable
- Implement proper concurrency where possible
- Avoid blocking operations
- Handle cancellation gracefully

Optimization Guidelines:
- Use batch operations for multiple cells
- Optimize range operations
- Minimize XML generation overhead
- Efficient image processing

## Maintenance & Updates

Backward Compatibility:
- Maintain API compatibility
- Use deprecation warnings for changes
- Provide migration guides
- Version APIs appropriately

Code Quality:
- Regular code reviews
- Automated testing
- Performance monitoring
- Documentation updates

Future Considerations:
- Plan for feature additions
- Consider platform expansion
- Monitor Swift evolution
- Track Excel format changes

Integration Guidelines:
- Use Swift Package Manager
- Maintain proper dependencies
- Version modules appropriately
- Document requirements

CI/CD Integration:
- Automated testing on macOS
- Code formatting checks
- Documentation generation
- Release automation

External Dependencies:
- Minimize external dependencies
- Use only essential libraries
- Document dependency reasons
- Monitor for updates

Troubleshooting:
- Sendable conformance warnings
- Memory usage with large files
- Image format detection failures
- CSV parsing edge cases
- Use proper logging
- Test with minimal examples
- Check file permissions
- Validate input data 