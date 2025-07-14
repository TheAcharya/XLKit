<p align="center">
  <a href="https://github.com/TheAcharya/XLKit"><img src="Assets/XLKit_Icon.png" height="200">
  <h1 align="center">XLKit</h1>
</p>


<p align="center"><a href="https://github.com/TheAcharya/XLKit/blob/main/LICENSE"><img src="http://img.shields.io/badge/license-MIT-lightgrey.svg?style=flat" alt="license"/></a>&nbsp;<a href="https://github.com/TheAcharya/XLKit"><img src="https://img.shields.io/badge/platform-macOS%20%7C%20iOS-lightgrey.svg?style=flat" alt="platform"/></a>&nbsp;<a href="https://github.com/TheAcharya/XLKit/actions/workflows/build.yml"><img src="https://github.com/TheAcharya/XLKit/actions/workflows/build.yml/badge.svg" alt="build"/></a></p>

A modern, ultra-easy Swift library for creating and manipulating Excel (.xlsx) files on macOS and iOS. XLKit provides a fluent, chainable API that makes Excel file generation effortless while supporting advanced features like image embedding, CSV/TSV import/export, cell formatting, and both synchronous and asynchronous operations. Built with Swift 6.0 and targeting macOS 12+ and iOS 15+, it offers type-safe operations, comprehensive error handling, and security features. Note: iOS support is available but not tested.

Purpose-built for [MarkersExtractor](https://github.com/TheAcharya/MarkersExtractor) - a tool for extracting markers from Final Cut Pro FCPXML files and generating comprehensive Excel reports with embedded images, CSV/TSV manifests, and structured data exports. Perfect for professional video editing workflows requiring pixel-perfect image embedding with all video and cinema aspect ratios.

Includes comprehensive security features for production use: rate limiting, security logging, file quarantine, input validation, and optional checksum verification to protect against vulnerabilities and supply chain attacks.

This codebase is developed using AI agents.

## Table of Contents

- [Features](#features)
- [Security Features](#security-features)
- [Performance Considerations](#performance-considerations)
- [Requirements](#requirements)
- [File Format](#file-format)
- [Quick Start](#quick-start)
- [Core Concepts](#core-concepts)
- [Basic Usage](#basic-usage)
- [CSV/TSV Import & Export](#csvtsv-import--export)
- [Image Support](#image-support)
- [Advanced Usage](#advanced-usage)
- [Error Handling](#error-handling)
- [Testing & Validation](#testing--validation)
- [Code Style & Formatting](#code-style--formatting)
- [Credits](#credits)
- [License](#License)
- [Reporting Bugs](#reporting-bugs)
- [Contribution](#contribution)

## Features

- Effortless API: Fluent, chainable, and bulk data helpers
- Perfect Image Embedding: Pixel-perfect image embedding with automatic aspect ratio preservation
- Professional Video Support: All 17 video and cinema aspect ratios with zero distortion
- Auto Cell Sizing: Automatic column width and row height calculation to perfectly fit images
- CSV/TSV Import/Export: Built-in support for importing and exporting CSV/TSV data
- Async & Sync Operations: Save workbooks with one line (async or sync)
- Type-Safe: Strong enums and structs for all data types
- Excel Compliance: Full OpenXML compliance with CoreXLSX validation
- No Dependencies: Pure Swift, macOS 12+, Swift 6.0+
- Comprehensive Testing: 46 tests with 100% API coverage and automated validation
- Security Features: Comprehensive security features for production use

## Security Features

XLKit includes comprehensive security features to protect against vulnerabilities, supply chain attacks, and malicious code injection:

### SecurityManager

Centralized security controls and monitoring with configurable features:

```swift
// Security features are automatically active
// Checksum storage can be enabled if needed
SecurityManager.enableChecksumStorage = true
```

### Rate Limiting

Prevents abuse and resource exhaustion:
- Default: 100 operations per minute
- Configurable: Time window and operation limits
- Automatic: Integrated into all file operations
- Logging: Rate limit violations are logged

### Security Logging

Comprehensive audit trail of all security-relevant operations:
- Console logging: Real-time security events
- File logging: Persistent audit trail
- Structured data: Timestamp, operation, details, user agent
- Operations logged: File generation, checksums, quarantines, rate limits

### File Quarantine

Isolates suspicious files to prevent execution:
- Pattern detection: Checks for malicious code patterns
- Size validation: Prevents oversized files
- Format validation: Validates image formats
- Automatic isolation: Moves suspicious files to quarantine directory

### File Checksums

Cryptographic integrity verification (optional):
- SHA-256 hashes: Secure file integrity verification
- Configurable: Can be enabled/disabled via `enableChecksumStorage`
- Tamper detection: Identifies unauthorized file modifications
- Supply chain protection: Ensures file authenticity

### Input Validation

Comprehensive validation of all user inputs:
- File paths: Validates and sanitizes file paths
- Image data: Validates image formats and sizes
- CSV data: Validates CSV structure and content
- Coordinates: Validates Excel coordinate formats

### Security Integration

Security features are integrated throughout the codebase:
- XLSXEngine: Rate limiting, logging, checksums for file generation
- ImageUtils: Quarantine, validation for image processing
- XLKit API: Input validation, security logging for all operations
- Test Runner: Security validation for all test operations

## Performance Considerations

- Memory Usage: XLKit is optimized for large datasets with efficient memory management
- Async Operations: Use async/await for file operations to avoid blocking the main thread
- Batch Operations: Set multiple cells in batches for better performance
- Range Operations: Use `setRange()` for setting multiple cells with the same value

## File Format

- Microsoft Excel
- Google Sheets
- LibreOffice Calc
- Numbers (macOS)
- Any application that supports the OpenXML format

## Requirements

- macOS: 12.0+
- iOS: 15.0+ (available but not tested)
- Swift: 6.0+

## Quick Start

```swift
import XLKit

// 1. Create a workbook and sheet
let workbook = XLKit.createWorkbook()
let sheet = workbook.addSheet(name: "Employees")

// 2. Add headers and data (fluent, chainable)
sheet
    .setRow(1, values: [.string("Name"), .string("Photo"), .string("Age")])
    .setRow(2, values: [.string("Alice"), .empty, .number(30)])
    .setRow(3, values: [.string("Bob"), .empty, .number(28)])

// 3. Add a GIF image to a cell with perfect aspect ratio preservation
let gifData = try Data(contentsOf: URL(fileURLWithPath: "alice.gif"))
try sheet.embedImageAutoSized(gifData, at: "B2", workbook: workbook)

// 4. Save the workbook (sync or async)
try XLKit.saveWorkbook(workbook, to: URL(fileURLWithPath: "employees.xlsx"))
// or
// try await XLKit.saveWorkbook(workbook, to: url)
```

## Core Concepts

### Workbook
A workbook contains multiple sheets and represents the entire Excel file.

```swift
let workbook = XLKit.createWorkbook()

// Add sheets
let sheet1 = workbook.addSheet(name: "Sheet1")
let sheet2 = workbook.addSheet(name: "Sheet2")

// Get sheets
let allSheets = workbook.getSheets()
let specificSheet = workbook.getSheet(name: "Sheet1")

// Remove sheets
workbook.removeSheet(name: "Sheet1")
```

### Sheet
A sheet represents a worksheet within the workbook.

```swift
let sheet = workbook.addSheet(name: "Data")

// Set cell values
sheet.setCell("A1", value: .string("Hello"))
sheet.setCell("B1", value: .number(42.5))
sheet.setCell("C1", value: .integer(100))
sheet.setCell("D1", value: .boolean(true))
sheet.setCell("E1", value: .date(Date()))
sheet.setCell("F1", value: .formula("=A1+B1"))

// Get cell values
let value = sheet.getCell("A1")

// Set cells by row/column
sheet.setCell(row: 1, column: 1, value: .string("A1"))

// Set ranges
sheet.setRange("A1:C3", value: .string("Range"))

// Merge cells
sheet.mergeCells("A1:C1")

// Get used cells
let usedCells = sheet.getUsedCells()

// Clear all data
sheet.clear()
```

### Cell Values

XLKit supports various cell value types:

```swift
// String values
.string("Hello World")

// Numeric values
.number(42.5)      // Double
.integer(100)      // Int

// Boolean values
.boolean(true)
.boolean(false)

// Date values
.date(Date())

// Formulas
.formula("=A1+B1")
.formula("=SUM(A1:A10)")

// Empty cells
.empty
```

### Cell Coordinates

Work with Excel-style coordinates:

```swift
// Create coordinates
let coord = CellCoordinate(row: 1, column: 1)
print(coord.excelAddress) // "A1"

// Parse Excel addresses
let coord2 = CellCoordinate(excelAddress: "B2")
print(coord2?.row)    // 2
print(coord2?.column) // 2

// Create ranges
let range = CellRange(start: coord, end: CellCoordinate(row: 3, column: 3))
print(range.excelRange) // "A1:C3"

// Parse Excel ranges
let range2 = CellRange(excelRange: "A1:B5")
```

## Basic Usage

### API Highlights

- Workbook: `createWorkbook()`, `addSheet(name:)`, `save(to:)`
- Sheet: `setCell`, `setRow`, `setColumn`, `setRange`, `mergeCells`, `embedImageAutoSized`, `setColumnWidth`
- Images: GIF, PNG, JPEG with perfect aspect ratio preservation
- CSV/TSV: `createWorkbookFromCSV`, `exportSheetToCSV`, `importCSVIntoSheet`
- Fluent API: Most setters return `Self` for chaining
- Bulk Data: `setRow`, `setColumn` for easy import
- Doc Comments: All public APIs are documented for Xcode autocomplete

### Example: Bulk Data and Images

```swift
let sheet = workbook.addSheet(name: "Products")
    .setRow(1, values: [.string("Product"), .string("Image"), .string("Price")])
    .setRow(2, values: [.string("Apple"), .empty, .number(1.99)])
    .setRow(3, values: [.string("Banana"), .empty, .number(0.99)])

let appleGif = try Data(contentsOf: URL(fileURLWithPath: "apple.gif"))
try sheet.embedImageAutoSized(appleGif, at: "B2", workbook: workbook)
```

### Cell Sizing

```swift
sheet.setColumnWidth(2, width: 200) // Set manually
// Automatic sizing is handled by embedImageAutoSized for perfect fit
```

## CSV/TSV Import & Export

XLKit provides simple static methods for importing and exporting CSV/TSV data:

```swift
// Create a workbook from CSV
let csvData = """
Name,Age,Salary
John,30,50000.5
Jane,25,45000.75
"""
let workbook = XLKit.createWorkbookFromCSV(csvData: csvData, hasHeader: true)
let sheet = workbook.getSheets().first!

// Export a sheet to CSV
let csv = XLKit.exportSheetToCSV(sheet: sheet)

// Import CSV into an existing sheet
XLKit.importCSVIntoSheet(sheet: sheet, csvData: csvData, hasHeader: true)

// Create a workbook from TSV
let tsvData = """
Product\tPrice\tIn Stock
Apple\t1.99\ttrue
Banana\t0.99\tfalse
"""
let tsvWorkbook = XLKit.createWorkbookFromTSV(tsvData: tsvData, hasHeader: true)
let tsvSheet = tsvWorkbook.getSheets().first!

// Export a sheet to TSV
let tsv = XLKit.exportSheetToTSV(sheet: tsvSheet)

// Import TSV into an existing sheet
XLKit.importTSVIntoSheet(sheet: tsvSheet, tsvData: tsvData, hasHeader: true)
```

All CSV/TSV helpers are available as static methods on `XLKit` for convenience, and are powered by the `XLKitFormatters` module under the hood.

## Image Support

### Perfect Aspect Ratio Preservation

XLKit provides pixel-perfect image embedding with automatic aspect ratio preservation. Images maintain their exact original proportions regardless of cell dimensions, preventing any stretching, squashing, or distortion.

Comprehensive Aspect Ratio Support: XLKit has been extensively tested with all professional video and cinema aspect ratios including:
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

All aspect ratios are preserved with pixel-perfect accuracy using empirically derived Excel formulas, ensuring professional-quality exports for any video or cinema format.

### Supported Image Formats
- GIF (including animated)
- PNG
- JPEG/JPG

### Adding Images with Perfect Sizing

```swift
// Add image with automatic cell sizing and aspect ratio preservation
let imageData = try Data(contentsOf: URL(fileURLWithPath: "image.png"))
try sheet.embedImageAutoSized(imageData, at: "A1", workbook: workbook)

// Add image from URL with perfect aspect ratio
let imageURL = URL(fileURLWithPath: "image.gif")
let imageData = try Data(contentsOf: imageURL)
try sheet.embedImageAutoSized(imageData, at: "B1", workbook: workbook)

// XLKit convenience method with scaling options
try XLKit.embedImage(
    imageData,
    at: "C1",
    in: sheet,
    of: workbook,
    scale: 1.0,
    maxWidth: 600,
    maxHeight: 400
)
```

### Image Scaling API

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

### Automatic Cell Sizing

The `embedImageAutoSized` method automatically:
- Calculates the perfect column width and row height for the image
- Preserves the exact aspect ratio of the original image
- Positions the image precisely at cell boundaries with zero offsets
- Uses empirically derived Excel formulas for pixel-perfect sizing

## Advanced Usage

### Multiple Sheets with Formulas

```swift
let workbook = XLKit.createWorkbook()

// Data sheet
let dataSheet = workbook.addSheet(name: "Data")
dataSheet.setCell("A1", value: .string("Product"))
dataSheet.setCell("B1", value: .string("Price"))
dataSheet.setCell("C1", value: .string("Quantity"))

dataSheet.setCell("A2", value: .string("Apple"))
dataSheet.setCell("B2", value: .number(1.99))
dataSheet.setCell("C2", value: .integer(10))

dataSheet.setCell("A3", value: .string("Orange"))
dataSheet.setCell("B3", value: .number(2.49))
dataSheet.setCell("C3", value: .integer(5))

// Summary sheet with formulas
let summarySheet = workbook.addSheet(name: "Summary")
summarySheet.setCell("A1", value: .string("Total Items"))
summarySheet.setCell("B1", value: .formula("=SUM(Data!C:C)"))

summarySheet.setCell("A2", value: .string("Total Revenue"))
summarySheet.setCell("B2", value: .formula("=SUMPRODUCT(Data!B:B,Data!C:C)"))

summarySheet.setCell("A3", value: .string("Average Price"))
summarySheet.setCell("B3", value: .formula("=AVERAGE(Data!B:B)"))

// Merge headers
dataSheet.mergeCells("A1:C1")
summarySheet.mergeCells("A1:B1")
```

### Working with Ranges

```swift
let sheet = workbook.addSheet(name: "Range Test")

// Set values in a range
sheet.setRange("A1:C3", value: .string("Range"))

// Create and work with ranges
let range = CellRange(excelRange: "A1:B5")!
for coordinate in range.coordinates {
    sheet.setCell(coordinate.excelAddress, value: .string("Cell \(coordinate.excelAddress)"))
}

// Merge multiple ranges
sheet.mergeCells("A1:C1")
sheet.mergeCells("A5:C5")
```

### Cell Formatting

```swift
let sheet = workbook.addSheet(name: "Formatted")

// Use predefined formats
sheet.setCell("A1", cell: Cell.string("Header", format: CellFormat.header()))
sheet.setCell("B1", cell: Cell.number(1234.56, format: CellFormat.currency()))
sheet.setCell("C1", cell: Cell.number(0.85, format: CellFormat.percentage()))

// Custom formatting
let customFormat = CellFormat(
    fontName: "Arial",
    fontSize: 14.0,
    fontWeight: .bold,
    backgroundColor: "#E0E0E0",
    horizontalAlignment: .center
)
sheet.setCell("D1", cell: Cell.string("Custom", format: customFormat))
```

## Error Handling

XLKit provides comprehensive error handling with specific error types:

```swift
do {
    try await XLKit.saveWorkbook(workbook, to: fileURL)
} catch XLKitError.invalidCoordinate(let coord) {
    print("Invalid coordinate: \(coord)")
} catch XLKitError.fileWriteError(let message) {
    print("File write error: \(message)")
} catch XLKitError.zipCreationError(let message) {
    print("ZIP creation error: \(message)")
} catch XLKitError.imageProcessingError(let message) {
    print("Image processing error: \(message)")
} catch XLKitError.csvParsingError(let message) {
    print("CSV parsing error: \(message)")
} catch {
    print("Unknown error: \(error)")
}
```

## Testing & Validation

XLKit includes comprehensive testing and validation capabilities with integrated security features:

### Unit Tests

The library includes 46 comprehensive unit tests covering:
- Core Workbook & Sheet Tests: Creation, management, and operations
- Cell Operations & Data Types: All cell value types and operations
- Coordinate & Range Tests: Excel coordinate parsing and range operations
- Image & Aspect Ratio Tests: All 17 professional aspect ratios with perfect preservation
- CSV/TSV Import/Export: Complete import/export functionality
- Cell Formatting: Predefined and custom formatting options
- Column & Row Sizing: Automatic sizing and manual adjustments
- File Operations: Async and sync workbook saving
- Error Handling: Comprehensive error testing and edge cases

### XLKitTestRunner

A modular command-line tool for generating Excel files for testing and demonstration:

```bash
# Run specific test types
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner security-demo
swift run XLKitTestRunner help
```

Available Test Types:
- `no-embeds` / `no-images`: Generate Excel from CSV without images
- `embed` / `with-embeds` / `with-images`: Generate Excel with embedded images from CSV data
- `comprehensive` / `demo`: Comprehensive API demonstration with all features
- `security-demo` / `security`: Demonstrate file path security restrictions
- `help` / `-h` / `--help`: Show available commands

Test Features:
- Security Integration: All tests include security logging and validation
- CoreXLSX Validation: Generated files are validated for Excel compliance
- Aspect Ratio Testing: Image embedding tests all 17 professional aspect ratios
- Performance Testing: Large dataset handling and memory optimization
- Error Handling: Comprehensive error testing and edge case coverage

### Security Features in Tests

All test operations include comprehensive security validation:
- Rate Limiting: Prevents test abuse and resource exhaustion
- Security Logging: All test operations are logged for audit trails
- Input Validation: All test inputs are validated for security
- File Quarantine: Suspicious test files are automatically quarantined
- Checksum Verification: Optional file integrity verification (disabled by default)

### Automated Validation

Every generated Excel file is automatically validated using CoreXLSX to ensure:
- Full OpenXML compliance and Excel compatibility
- Proper file structure and relationships
- Image embedding integrity with perfect aspect ratio preservation
- Cell and row data accuracy
- Professional-quality exports for all video and cinema formats
- Zero tolerance for distortion or stretching in embedded images
- Security compliance and audit trail integrity

### Test Output Structure

Generated test files are saved to:
```
Test-Workflows/
├── Embed-Test.xlsx          # From no-embeds test
├── Embed-Test-Embed.xlsx    # From embed test (with images)
├── Comprehensive-Demo.xlsx  # From comprehensive test
└── [Your-Test].xlsx         # From custom tests
```

### Test Coverage

- Unit Tests: All public APIs and core functionality
- Integration Tests: Complete Excel file generation workflows
- Image Embedding Tests: Aspect ratio preservation for all 17 supported formats
- CSV/TSV Tests: Import/export functionality with various data types
- Performance Tests: Large dataset handling and memory management
- Validation Tests: CoreXLSX compliance verification for all generated files
- Security Tests: Rate limiting, input validation, file quarantine, and checksum verification

## Code Style & Formatting

XLKit enforces a modern, consistent Swift code style for all contributions:

- 4-space indentation
- Trailing commas
- Grouped and reordered imports
- 120 character line length
- Consistent spacing and blank lines
- No force-unwraps or force-casts in public API
- All public APIs have doc comments
- Follows Swift 6 idioms and best practices

A `.swift-format` file is included in the repo. To format the codebase, run:

```sh
swift-format format -i .
```

Or use Xcode's built-in formatter for most style rules.

All code must be formatted and pass CI before merging. See `.cursorrules` for more details.

## Credits

Created by [Vigneswaran Rajkumar](https://bsky.app/profile/vigneswaranrajkumar.com)

## License

Licensed under the MIT license. See [LICENSE](https://github.com/TheAcharya/XLKit/blob/main/LICENSE) for details.

## Reporting Bugs

For bug reports, feature requests and suggestions you can create a new [issue](https://github.com/TheAcharya/XLKit/issues) to discuss.

## Contribution

Community contributions are welcome and appreciated. Developers are encouraged to fork the repository and submit pull requests to enhance functionality or introduce thoughtful improvements. However, a key requirement is that nothing should break—all existing features and behaviours and logic must remain fully functional and unchanged. Once reviewed and approved, updates will be merged into the main branch.