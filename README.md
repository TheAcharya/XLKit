<p align="center">
  <a href="https://github.com/TheAcharya/XLKit"><img src="Assets/XLKit_Icon.png" height="200">
  <h1 align="center">XLKit</h1>
</p>


<p align="center"><a href="https://github.com/TheAcharya/XLKit/blob/main/LICENSE"><img src="http://img.shields.io/badge/license-MIT-lightgrey.svg?style=flat" alt="license"/></a>&nbsp;<a href="https://github.com/TheAcharya/XLKit"><img src="https://img.shields.io/badge/platform-macOS%20%7C%20iOS-lightgrey.svg?style=flat" alt="platform"/></a>&nbsp;<a href="https://github.com/TheAcharya/XLKit/actions/workflows/build.yml"><img src="https://github.com/TheAcharya/XLKit/actions/workflows/build.yml/badge.svg" alt="build"/></a>&nbsp;<a href="https://github.com/TheAcharya/XLKit/actions/workflows/codeql.yml"><img src="https://github.com/TheAcharya/XLKit/actions/workflows/codeql.yml/badge.svg" alt="codeql"/></a></p>

A modern Swift library for creating and manipulating Excel (.xlsx) files on macOS and iOS. XLKit provides a fluent, chainable API that makes Excel file generation effortless while supporting advanced features like image embedding, CSV/TSV import/export, cell formatting, and both synchronous and asynchronous operations. Built with Swift 6.0 and targeting macOS 12+ and iOS 15+, it offers type-safe operations, comprehensive error handling, and security features. iOS support is available and tested in CI/CD.

Purpose-built for [MarkersExtractor](https://github.com/TheAcharya/MarkersExtractor) - a tool for extracting markers from Final Cut Pro FCPXML files and generating comprehensive Excel reports with embedded images, CSV/TSV manifests, and structured data exports. Perfect for professional video editing workflows requiring pixel-perfect image embedding with all video and cinema aspect ratios.

Includes comprehensive security features for production use: rate limiting, security logging, file quarantine, input validation, and optional checksum verification to protect against vulnerabilities and supply chain attacks.

This codebase is developed using AI agents.

> [!IMPORTANT]
> XLKit has yet to be extensively tested in production environments, real-world workflows, or application integration. This library serves as a modernised foundation for AI-assisted development and experimentation with Excel file generation, and image embedding. While comprehensive testing has been performed, production deployment should be approached with appropriate caution and validation. Additionally, this project would not be actively maintained, so please consider this when planning long-term integrations.

## Table of Contents

- [Features](#features)
- [Security Features](#security-features)
  - [SecurityManager](#securitymanager)
  - [Rate Limiting](#rate-limiting)
  - [Security Logging](#security-logging)
  - [File Quarantine](#file-quarantine)
  - [File Checksums](#file-checksums)
  - [Input Validation](#input-validation)
  - [Security Integration](#security-integration)
- [Performance Considerations](#performance-considerations)
- [Requirements](#requirements)
- [iOS Support](#ios-support)
  - [iOS File System Considerations](#ios-file-system-considerations)
  - [iOS Configuration Requirements](#ios-configuration-requirements)
  - [iOS-Specific Features](#ios-specific-features)
  - [iOS Testing](#ios-testing)
- [Installing](#installing)
  - [Swift Package Manager](#swift-package-manager)
  - [Import](#import)
  - [Verify Installation](#verify-installation)
- [File Format](#file-format)
- [Quick Start](#quick-start)
- [Core Concepts](#core-concepts)
  - [Workbook](#workbook)
  - [Sheet](#sheet)
  - [Cell Values](#cell-values)
  - [Cell Coordinates](#cell-coordinates)
- [Basic Usage](#basic-usage)
  - [API Highlights](#api-highlights)
  - [Example: Bulk Data and Images](#example-bulk-data-and-images)
  - [Cell Sizing](#cell-sizing)
  - [Utility Properties and Methods](#utility-properties-and-methods)
- [CSV/TSV Import & Export](#csvtsv-import--export)
- [Image Support](#image-support)
  - [Perfect Aspect Ratio Preservation](#perfect-aspect-ratio-preservation)
  - [Supported Image Formats](#supported-image-formats)
  - [Adding Images with Perfect Sizing](#adding-images-with-perfect-sizing)
  - [Image Scaling API](#image-scaling-api)
  - [Automatic Cell Sizing](#automatic-cell-sizing)
- [Number Format Support](#number-format-support)
  - [Supported Number Formats](#supported-number-formats)
  - [Number Format Examples](#number-format-examples)
  - [Font Colour with Number Formats](#font-colour-with-number-formats)
  - [Excel Compliance](#excel-compliance)
  - [Testing Number Formats](#testing-number-formats)
- [Text Alignment Support](#text-alignment-support)
  - [Horizontal Alignment (5 options)](#horizontal-alignment-5-options)
  - [Vertical Alignment (5 options)](#vertical-alignment-5-options)
  - [Combined Alignment](#combined-alignment)
  - [Practical Examples](#practical-examples)
  - [Predefined Formats with Alignment](#predefined-formats-with-alignment)
  - [Alignment with Other Formatting](#alignment-with-other-formatting)
- [Advanced Usage](#advanced-usage)
  - [Multiple Sheets with Formulas](#multiple-sheets-with-formulas)
  - [Working with Ranges](#working-with-ranges)
  - [Cell Formatting](#cell-formatting)
  - [Font Colour Support](#font-colour-support)
- [Error Handling](#error-handling)
- [Testing & Validation](#testing--validation)
  - [Unit Tests](#unit-tests)
  - [XLKitTestRunner](#xlkittestrunner)
  - [iOS Compatibility Testing](#ios-compatibility-testing)
  - [Security Features in Tests](#security-features-in-tests)
  - [Automated Validation](#automated-validation)
  - [Test Output Structure](#test-output-structure)
  - [CI/CD Integration](#cicd-integration)
  - [Test Coverage](#test-coverage)
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
- Cell Formatting: Comprehensive formatting including font colours, backgrounds, borders, and text alignment (all 6 alignment options) with proper XML generation
- CSV/TSV Import/Export: Built-in support for importing and exporting CSV/TSV data
- Async & Sync Operations: Save workbooks with one line (async or sync)
- Type-Safe: Strong enums and structs for all data types
- Excel Compliance: Full OpenXML compliance with CoreXLSX validation
- No Dependencies: Pure Swift, macOS 12+, Swift 6.0+
- Comprehensive Testing: 45 tests with 100% API coverage, including all text alignment options (horizontal, vertical, combined), number formatting, and automated validation
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

## iOS Support

XLKit is fully supported on iOS 15+ with platform-specific optimizations for iOS sandbox restrictions.

### iOS File System Considerations

On iOS, apps run in a sandboxed environment with restricted file system access. XLKit automatically handles these restrictions:

```swift
import XLKit

// Recommended: Use CoreUtils.safeFileURL() for iOS
let safeURL = CoreUtils.safeFileURL(for: "employees.xlsx")
try await workbook.save(to: safeURL)

// Also works: Use documents directory directly
if let documentsURL = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask).first {
    let fileURL = documentsURL.appendingPathComponent("employees.xlsx")
    try await workbook.save(to: fileURL)
}

// Avoid: Using arbitrary file paths on iOS
// try await workbook.save(to: URL(fileURLWithPath: "employees.xlsx")) // May fail
```

#### iOS Configuration Requirements

To enable document browser access and allow users to load Excel files from the Files app on iOS, add the following to your app's Info.plist or target configuration:

**Info.plist:**
```xml
<key>UIFileSharingEnabled</key>
<true/>
<key>LSSupportsOpeningDocumentsInPlace</key>
<true/>
```

**Or in Xcode target settings:**
- Set "Supports Document Browser" to `YES`
- Set "Application supports iTunes file sharing" to `YES`

This configuration allows users to:
- Access Excel files in the Files app
- Import Excel files from other apps
- Share Excel files through the system share sheet
- Use the document browser to manage Excel files

### iOS-Specific Features

- Automatic sandbox compliance: XLKit automatically uses iOS-appropriate directories
- Copy-based file operations: Uses copy instead of move for better iOS compatibility
- Relaxed path validation: iOS sandbox handles actual restrictions
- Documents directory support: Automatically includes documents and caches directories
- Temporary directory fallback: Falls back to temporary directory if needed

### iOS Testing

XLKit is tested on iOS in CI/CD with the following workflow:
- Build verification on iOS Simulator
- Unit test execution on iOS
- Image embedding tests with perfect aspect ratio preservation
- File system operations in iOS sandbox environment

## Installing

### Swift Package Manager

XLKit is available through Swift Package Manager. Add it to your project dependencies:

#### Xcode
1. In Xcode, go to **File** → **Add Package Dependencies**
2. Enter the repository URL: `https://github.com/TheAcharya/XLKit.git`
3. Click **Add Package**
4. Select the XLKit product and click **Add Package**

#### Package.swift
Add XLKit to your `Package.swift` dependencies:

```swift
dependencies: [
    .package(url: "https://github.com/TheAcharya/XLKit.git", from: "1.0.7")
]
```

#### Command Line
```bash
swift package add https://github.com/TheAcharya/XLKit.git
```

### Import

Once installed, import XLKit in your Swift files:

```swift
import XLKit
```

### Verify Installation

Test that XLKit is working correctly:

```swift
import XLKit

// Create a simple workbook
let workbook = Workbook()
let sheet = workbook.addSheet(name: "Test")
sheet.setCell("A1", value: .string("Hello, XLKit!"))

// Save to verify everything works
let safeURL = CoreUtils.safeFileURL(for: "test.xlsx")
try workbook.save(to: safeURL)
print("XLKit is working correctly!")
```

## Quick Start

```swift
import XLKit

// 1. Create a workbook and sheet
let workbook = Workbook()
let sheet = workbook.addSheet(name: "Employees")

// 2. Add headers and data (fluent, chainable)
sheet
    .setRow(1, strings: ["Name", "Photo", "Age"])
    .setRow(2, strings: ["Alice", "", "30"])
    .setRow(3, strings: ["Bob", "", "28"])

// 3. Add formatting with font colours
sheet.setCell("A1", string: "Name", format: CellFormat.header())
sheet.setCell("B1", string: "Photo", format: CellFormat.header())
sheet.setCell("C1", string: "Age", format: CellFormat.coloredText(color: "#FF0000"))

// 4. Add a GIF image to a cell with perfect aspect ratio preservation
let gifData = try Data(contentsOf: URL(fileURLWithPath: "alice.gif"))
try await sheet.embedImageAutoSized(gifData, at: "B2", of: workbook)

// 5. Save the workbook (sync or async)
try await workbook.save(to: URL(fileURLWithPath: "employees.xlsx"))
// or
// try workbook.save(to: url)

// iOS Note: For iOS apps, use CoreUtils.safeFileURL() to get a safe file path:
// let safeURL = CoreUtils.safeFileURL(for: "employees.xlsx")
// try await workbook.save(to: safeURL)
```

## Core Concepts

### Workbook
A workbook contains multiple sheets and represents the entire Excel file.

```swift
let workbook = Workbook()

// Add sheets
let sheet1 = workbook.addSheet(name: "Sheet1")
let sheet2 = workbook.addSheet(name: "Sheet2")

// Get sheets
let allSheets = workbook.getSheets()
let specificSheet = workbook.getSheet(name: "Sheet1")

// Remove sheets
workbook.removeSheet(name: "Sheet1")

// Image management
workbook.addImage(excelImage)
let image = workbook.getImage(withId: "image-id")
workbook.removeImage(withId: "image-id")
let pngImages = workbook.getImages(withFormat: .png)
workbook.clearImages()
let imageCount = workbook.imageCount
```

### Sheet
A sheet represents a worksheet within the workbook.

```swift
let sheet = workbook.addSheet(name: "Data")

// Set cell values (basic method)
sheet.setCell("A1", value: .string("Hello"))
sheet.setCell("B1", value: .number(42.5))
sheet.setCell("C1", value: .integer(100))
sheet.setCell("D1", value: .boolean(true))
sheet.setCell("E1", value: .date(Date()))
sheet.setCell("F1", value: .formula("=A1+B1"))

// Set cell values with convenience methods (recommended)
sheet.setCell("A1", string: "Hello", format: CellFormat.header())
sheet.setCell("B1", number: 42.5, format: CellFormat.currency())
sheet.setCell("C1", integer: 100, format: CellFormat.number())
sheet.setCell("D1", boolean: true, format: CellFormat.text())
sheet.setCell("E1", date: Date(), format: CellFormat.date())
sheet.setCell("F1", formula: "=A1+B1", format: CellFormat.text())

// Get cell values
let value = sheet.getCell("A1")
let cellWithFormat = sheet.getCellWithFormat("A1")

// Set cells by row/column
sheet.setCell(row: 1, column: 1, value: .string("A1"))
sheet.setCell(row: 1, column: 1, cell: Cell.string("A1", format: CellFormat.header()))

// Set ranges (basic method)
sheet.setRange("A1:C3", value: .string("Range"))

// Set ranges with convenience methods (recommended)
sheet.setRange("A1:C3", string: "Range", format: CellFormat.bordered())
sheet.setRange("D1:F3", number: 42.5, format: CellFormat.currency())
sheet.setRange("G1:I3", integer: 100, format: CellFormat.number())
sheet.setRange("J1:L3", boolean: true, format: CellFormat.text())
sheet.setRange("M1:O3", date: Date(), format: CellFormat.date())
sheet.setRange("P1:R3", formula: "=A1+B1", format: CellFormat.text())

// Merge cells
sheet.mergeCells("A1:C1")

// Get used cells and ranges
let usedCells = sheet.getUsedCells()
let mergedRanges = sheet.getMergedRanges()

// Clear all data
sheet.clear()

// Utility properties
let allCells = sheet.allCells                    // [String: CellValue]
let allFormattedCells = sheet.allFormattedCells  // [String: Cell]
let isEmpty = sheet.isEmpty                      // Bool
let cellCount = sheet.cellCount                  // Int
let imageCount = sheet.imageCount                // Int
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

- Workbook: `Workbook()`, `addSheet(name:)`, `save(to:)`, image management methods
- Sheet: `setCell`, `setRow`, `setColumn`, `setRange`, `mergeCells`, `embedImageAutoSized`, `setColumnWidth`
- Convenience Methods: Type-specific setters like `setCell(string:format:)`, `setRange(number:format:)`
- Images: GIF, PNG, JPEG with perfect aspect ratio preservation
- CSV/TSV: `Workbook(fromCSV:)`, `exportToCSV()`, `importCSVIntoSheet`
- Fluent API: Most setters return `Self` for chaining
- Bulk Data: `setRow`, `setColumn` for easy import
- Utility Properties: `allCells`, `allFormattedCells`, `isEmpty`, `cellCount`, `imageCount`
- Doc Comments: All public APIs are documented for Xcode autocomplete

### Example: Bulk Data and Images

```swift
let sheet = workbook.addSheet(name: "Products")
    .setRow(1, strings: ["Product", "Image", "Price"])
    .setRow(2, strings: ["Apple", "", "1.99"])
    .setRow(3, strings: ["Banana", "", "0.99"])

// Add image with perfect aspect ratio preservation
let appleGif = try Data(contentsOf: URL(fileURLWithPath: "apple.gif"))
try await sheet.embedImageAutoSized(appleGif, at: "B2", of: workbook)

// Alternative: Use convenience methods for better type safety
sheet.setCell("A1", string: "Product", format: CellFormat.header())
sheet.setCell("C1", string: "Price", format: CellFormat.header())
sheet.setCell("A2", string: "Apple")
sheet.setCell("C2", number: 1.99, format: CellFormat.currency())

// Check sheet properties
print("Sheet has \(sheet.cellCount) cells and \(sheet.imageCount) images")
print("Sheet is empty: \(sheet.isEmpty)")
```

### Cell Sizing

```swift
sheet.setColumnWidth(2, width: 200) // Set manually
sheet.setRowHeight(1, height: 25.0)  // Set row height manually

// Get current dimensions
let colWidth = sheet.getColumnWidth(2)
let rowHeight = sheet.getRowHeight(1)
let allColWidths = sheet.getColumnWidths()
let allRowHeights = sheet.getRowHeights()

// Automatic sizing is handled by embedImageAutoSized for perfect fit
sheet.autoSizeColumn(2, forImageAt: "B2") // Auto-size column for specific image
```

### Utility Properties and Methods

```swift
// Sheet utility properties
let allCells = sheet.allCells                    // [String: CellValue] - All cells with values
let allFormattedCells = sheet.allFormattedCells  // [String: Cell] - All cells with formatting
let isEmpty = sheet.isEmpty                      // Bool - Check if sheet has any content
let cellCount = sheet.cellCount                  // Int - Number of cells with values
let imageCount = sheet.imageCount                // Int - Number of images in sheet

// Workbook utility properties
let workbookImageCount = workbook.imageCount     // Int - Total images in workbook

// Get all images
let sheetImages = sheet.getImages()              // [String: ExcelImage] - All images in sheet
let workbookImages = workbook.getImages()        // [ExcelImage] - All images in workbook
let pngImages = workbook.getImages(withFormat: .png) // [ExcelImage] - Images by format

// Check image existence
let hasImage = sheet.hasImage(at: "A1")          // Bool - Check if cell has image
let image = sheet.getImage(at: "A1")             // ExcelImage? - Get image at coordinate
let workbookImage = workbook.getImage(withId: "image-id") // ExcelImage? - Get image by ID
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
let workbook = Workbook(fromCSV: csvData, hasHeader: true)
let sheet = workbook.getSheets().first!

// Export a sheet to CSV
let csv = sheet.exportToCSV()

// Import CSV into an existing sheet
sheet.importCSV(csvData, hasHeader: true)

// Create a workbook from TSV
let tsvData = """
Product\tPrice\tIn Stock
Apple\t1.99\ttrue
Banana\t0.99\tfalse
"""
let tsvWorkbook = Workbook(fromTSV: tsvData, hasHeader: true)
let tsvSheet = tsvWorkbook.getSheets().first!

// Export a sheet to TSV
let tsv = tsvSheet.exportToTSV()

// Import TSV into an existing sheet
tsvSheet.importTSV(tsvData, hasHeader: true)
```

All CSV/TSV helpers are available as instance methods on `Workbook` and `Sheet` classes for convenience, and are powered by the `XLKitFormatters` module under the hood.

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
try await sheet.embedImageAutoSized(imageData, at: "A1", of: workbook)

// Add image from URL with perfect aspect ratio
let imageURL = URL(fileURLWithPath: "image.gif")
let imageData = try Data(contentsOf: imageURL)
try await sheet.embedImageAutoSized(imageData, at: "B1", of: workbook)

// Add image with scaling control
try await sheet.embedImage(
    imageData,
    at: "C1",
    of: workbook,
    scale: 0.7,
    maxWidth: 600,
    maxHeight: 400
)

// Add image from file URL
try await sheet.embedImage(from: imageURL, at: "D1")

// Add image from file path
try await sheet.embedImage(from: "/path/to/image.png", at: "E1")

// Manual image management
let excelImage = ExcelImage(
    id: "my-image",
    data: imageData,
    format: .png,
    originalSize: CGSize(width: 100, height: 100),
    displaySize: CGSize(width: 50, height: 50)
)
sheet.addImage(excelImage, at: "F1")
workbook.addImage(excelImage)

// Check if cell has image
let hasImage = sheet.hasImage(at: "A1")
let image = sheet.getImage(at: "A1")
sheet.removeImage(at: "A1")
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
let workbook = Workbook()

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

// Set values in a range (basic method)
sheet.setRange("A1:C3", value: .string("Range"))

// Set values in ranges with convenience methods (recommended)
sheet.setRange("A1:C3", string: "Range", format: CellFormat.bordered())
sheet.setRange("D1:F3", number: 42.5, format: CellFormat.currency())
sheet.setRange("G1:I3", integer: 100, format: CellFormat.number())
sheet.setRange("J1:L3", boolean: true, format: CellFormat.text())
sheet.setRange("M1:O3", date: Date(), format: CellFormat.date())
sheet.setRange("P1:R3", formula: "=A1+B1", format: CellFormat.text())

// Create and work with ranges
let range = CellRange(excelRange: "A1:B5")!
for coordinate in range.coordinates {
    sheet.setCell(coordinate.excelAddress, value: .string("Cell \(coordinate.excelAddress)"))
}

// Merge multiple ranges
sheet.mergeCells("A1:C1")
sheet.mergeCells("A5:C5")

// Get merged ranges
let mergedRanges = sheet.getMergedRanges()
```

### Cell Formatting

```swift
let sheet = workbook.addSheet(name: "Formatted")

// Use predefined formats with convenience methods (recommended)
sheet.setCell("A1", string: "Header", format: CellFormat.header())
sheet.setCell("B1", number: 1234.56, format: CellFormat.currency())
sheet.setCell("C1", number: 0.85, format: CellFormat.percentage())
sheet.setCell("D1", date: Date(), format: CellFormat.date())
sheet.setCell("E1", boolean: true, format: CellFormat.text())

// Use Cell struct for more control
sheet.setCell("F1", cell: Cell.string("Header", format: CellFormat.header()))
sheet.setCell("G1", cell: Cell.number(1234.56, format: CellFormat.currency()))

// Custom formatting
let customFormat = CellFormat(
    fontName: "Arial",
    fontSize: 14.0,
    fontWeight: .bold,
    backgroundColor: "#E0E0E0",
    horizontalAlignment: .center
)
sheet.setCell("H1", string: "Custom", format: customFormat)

// Font colour formatting
let redTextFormat = CellFormat.coloredText(color: "#FF0000")
let blueTextFormat = CellFormat.coloredText(color: "#0000FF")
let currencyRedFormat = CellFormat.currency(format: .currencyWithDecimals, color: "#FF0000")

sheet.setCell("I1", string: "Red Text", format: redTextFormat)
sheet.setCell("J1", string: "Blue Text", format: blueTextFormat)
sheet.setCell("K1", number: 1234.56, format: currencyRedFormat)

// Get cell with formatting
let cellWithFormat = sheet.getCellWithFormat("A1")
let cellFormat = sheet.getCellFormat("A1")

### Font Colour Support

XLKit provides comprehensive font colour support with proper XML generation and theme colour support:

```swift
// Basic font colours
let redText = CellFormat.coloredText(color: "#FF0000")
let blueText = CellFormat.coloredText(color: "#0000FF")
let greenText = CellFormat.coloredText(color: "#00FF00")

// Font colours with other formatting
let boldRedText = CellFormat(
    fontColor: "#FF0000",
    fontWeight: .bold,
    fontSize: 14.0
)

let colouredCurrency = CellFormat.currency(
    format: .currencyWithDecimals,
    color: "#FF0000"  // Red currency values
)

// Apply font colours to cells
sheet.setCell("A1", string: "Red Header", format: redText)
sheet.setCell("B1", number: 1234.56, format: colouredCurrency)
sheet.setCell("C1", string: "Bold Blue", format: boldRedText)

// Font colours work with all cell types
sheet.setCell("D1", boolean: true, format: CellFormat.coloredText(color: "#FF6600"))
sheet.setCell("E1", date: Date(), format: CellFormat.coloredText(color: "#6600FF"))
```

Supported colour formats:
- Hex colours: `#FF0000`, `#00FF00`, `#0000FF`
- Theme colours: Excel's built-in theme colour system
- All colours are properly converted to Excel's internal format

### Number Format Support

XLKit provides comprehensive number formatting support with proper Excel compliance. All number formats are correctly applied in Excel with thousands grouping, currency symbols, and proper display in the "Format Cells" dialog.

### Supported Number Formats

```swift
// Currency formatting
let currencyFormat = CellFormat.currency()                    // $1,234.56
let customCurrency = CellFormat.currency(format: .custom("$#,##0.00"))  // Custom currency

// Percentage formatting  
let percentageFormat = CellFormat.percentage()                // 12.34%
let customPercentage = CellFormat.percentage(format: .custom("0.00%"))  // Custom percentage

// Date formatting
let dateFormat = CellFormat.date()                            // Standard date format
let customDate = CellFormat.date(format: .custom("mmmm dd, yyyy"))      // Custom date

// Custom number formats
let thousandsFormat = CellFormat()
thousandsFormat.numberFormat = .custom
thousandsFormat.customNumberFormat = "#,##0"                  // 1,234,567

let decimalFormat = CellFormat()
decimalFormat.numberFormat = .custom
decimalFormat.customNumberFormat = "0.000"                    // 3.142

let mixedFormat = CellFormat()
mixedFormat.numberFormat = .custom
mixedFormat.customNumberFormat = "$#,##0.00;($#,##0.00)"     // ($1,234.56) for negatives
```

### Number Format Examples

```swift
let sheet = workbook.addSheet(name: "Number Formats")

// Currency examples
sheet.setCell("A1", number: 1234.56, format: CellFormat.currency())
sheet.setCell("A2", number: 5678.90, format: CellFormat.currency(format: .custom("$#,##0.00")))

// Percentage examples
sheet.setCell("B1", number: 0.1234, format: CellFormat.percentage())
sheet.setCell("B2", number: 0.5678, format: CellFormat.percentage(format: .custom("0.00%")))

// Custom number formats
var thousandsFormat = CellFormat()
thousandsFormat.numberFormat = .custom
thousandsFormat.customNumberFormat = "#,##0"
sheet.setCell("C1", number: 1234567, format: thousandsFormat)

var decimalFormat = CellFormat()
decimalFormat.numberFormat = .custom
decimalFormat.customNumberFormat = "0.000"
sheet.setCell("C2", number: 3.14159, format: decimalFormat)

// Date formatting
sheet.setCell("D1", date: Date(), format: CellFormat.date())
var customDateFormat = CellFormat()
customDateFormat.numberFormat = .custom
customDateFormat.customNumberFormat = "mmmm dd, yyyy"
sheet.setCell("D2", date: Date(), format: customDateFormat)
```

### Font Colour with Number Formats

```swift
// Red currency values
let redCurrencyFormat = CellFormat.currency(
    format: .currencyWithDecimals,
    color: "#FF0000"
)

// Blue percentage values
let bluePercentageFormat = CellFormat.percentage(
    format: .percentageWithDecimals,
    color: "#0000FF"
)

// Apply coloured number formats
sheet.setCell("A1", number: 1234.56, format: redCurrencyFormat)
sheet.setCell("B1", number: 0.85, format: bluePercentageFormat)
```

### Excel Compliance

All number formats are properly implemented with:
- Thousands grouping: Numbers display with proper separators (1,234.56)
- Currency symbols: Dollar signs and other currency symbols display correctly
- Format Cells dialog: Excel shows the correct format instead of "General"
- Custom formats: User-defined number formats work as expected
- Negative numbers: Proper handling of negative values with parentheses or minus signs
- Locale support: Respects system locale settings for decimal separators

### Testing Number Formats

Use the XLKitTestRunner to validate number formatting:

```bash
swift run XLKitTestRunner number-formats
```

This generates `Number-Format-Test.xlsx` with comprehensive examples of all number format types.

## Text Alignment Support

XLKit provides comprehensive text alignment support with all 6 alignment options available in Excel:

#### Horizontal Alignment (5 options)
```swift
// Left alignment
let leftAligned = CellFormat(horizontalAlignment: .left)

// Center alignment
let centerAligned = CellFormat(horizontalAlignment: .center)

// Right alignment
let rightAligned = CellFormat(horizontalAlignment: .right)

// Justified alignment
let justified = CellFormat(horizontalAlignment: .justify)

// Distributed alignment
let distributed = CellFormat(horizontalAlignment: .distributed)
```

#### Vertical Alignment (5 options)
```swift
// Top alignment
let topAligned = CellFormat(verticalAlignment: .top)

// Center alignment
let centerVertically = CellFormat(verticalAlignment: .center)

// Bottom alignment
let bottomAligned = CellFormat(verticalAlignment: .bottom)

// Justified alignment
let justifiedVertically = CellFormat(verticalAlignment: .justify)

// Distributed alignment
let distributedVertically = CellFormat(verticalAlignment: .distributed)
```

#### Combined Alignment
```swift
// Center both horizontally and vertically
let centerCenter = CellFormat(
    horizontalAlignment: .center,
    verticalAlignment: .center
)

// Top-right alignment
let topRight = CellFormat(
    horizontalAlignment: .right,
    verticalAlignment: .top
)

// Bottom-left alignment
let bottomLeft = CellFormat(
    horizontalAlignment: .left,
    verticalAlignment: .bottom
)
```

#### Practical Examples
```swift
// Headers with center alignment
let headerFormat = CellFormat(
    fontWeight: .bold,
    fontSize: 14.0,
    backgroundColor: "#E6E6E6",
    horizontalAlignment: .center
)

// Right-aligned numbers for better readability
let numberFormat = CellFormat(
    numberFormat: .currencyWithDecimals,
    horizontalAlignment: .right
)

// Top-aligned text for multi-line content
let multiLineFormat = CellFormat(
    textWrapping: true,
    verticalAlignment: .top
)

// Apply alignment to cells
sheet.setCell("A1", string: "Centered Header", format: headerFormat)
sheet.setCell("B1", number: 1234.56, format: numberFormat)
sheet.setCell("C1", string: "Multi-line\ntext content", format: multiLineFormat)
```

#### Predefined Formats with Alignment
```swift
// Header format automatically centers text
let header = CellFormat.header(fontSize: 14.0, backgroundColor: "#E6E6E6")
// This sets horizontalAlignment = .center internally

// Currency format with right alignment
let currencyRight = CellFormat.currency(
    format: .currencyWithDecimals,
    color: "#FF0000"
)
// Add right alignment
currencyRight.horizontalAlignment = .right
```

#### Alignment with Other Formatting
```swift
// Bold, red text, center-aligned
let boldRedCenter = CellFormat(
    fontWeight: .bold,
    fontColor: "#FF0000",
    horizontalAlignment: .center,
    verticalAlignment: .center
)

// Bordered cells with distributed alignment
let borderedDistributed = CellFormat(
    borderTop: .thin,
    borderBottom: .thin,
    borderLeft: .thin,
    borderRight: .thin,
    horizontalAlignment: .distributed,
    verticalAlignment: .distributed
)
```

All alignment options are properly converted to Excel's OpenXML format and will display correctly in Excel, Google Sheets, LibreOffice Calc, and other spreadsheet applications.
```

## Error Handling

XLKit provides comprehensive error handling with specific error types:

```swift
do {
    try await workbook.save(to: fileURL)
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

The library includes 40 comprehensive unit tests covering:
- Core Workbook & Sheet Tests: Creation, management, and operations
- Cell Operations & Data Types: All cell value types and operations
- Coordinate & Range Tests: Excel coordinate parsing and range operations
- Image & Aspect Ratio Tests: All 17 professional aspect ratios with perfect preservation
- CSV/TSV Import/Export: Complete import/export functionality
- Cell Formatting: Predefined and custom formatting options including font colours and all text alignment options (horizontal, vertical, combined)
- Column & Row Sizing: Automatic sizing and manual adjustments
- File Operations: Async and sync workbook saving
- Error Handling: Comprehensive error testing and edge cases
- Platform Compatibility: iOS-specific file system operations and sandbox restrictions
- All text alignment options (horizontal, vertical, combined) are fully tested

### XLKitTestRunner

A modular command-line tool for generating Excel files for testing and demonstration:

```bash
# Run specific test types
swift run XLKitTestRunner no-embeds
swift run XLKitTestRunner embed
swift run XLKitTestRunner comprehensive
swift run XLKitTestRunner security-demo
swift run XLKitTestRunner ios-test
swift run XLKitTestRunner number-formats
swift run XLKitTestRunner help
```

Available Test Types:
- no-embeds / no-images: Generate Excel from CSV without images
- embed / with-embeds / with-images: Generate Excel with embedded images from CSV data
- comprehensive / demo: Comprehensive API demonstration with all features
- security-demo / security: Demonstrate file path security restrictions
- ios-test / ios: Test iOS file system compatibility and platform-specific features
- number-formats / formats: Test number formatting (currency, percentage, custom formats)
- help / -h / --help: Show available commands

Test Features:
- Security Integration: All tests include security logging and validation
- CoreXLSX Validation: Generated files are validated for Excel compliance
- Aspect Ratio Testing: Image embedding tests all 17 professional aspect ratios
- Performance Testing: Large dataset handling and memory optimisation
- Error Handling: Comprehensive error testing and edge case coverage
- Platform Testing: iOS compatibility validation and sandbox restrictions testing

### iOS Compatibility Testing

The `ios-test` command validates iOS compatibility by:
- Testing platform-specific file system operations
- Validating iOS sandbox restrictions handling
- Testing ZIP creation with iOS-compatible methods
- Ensuring security features work on iOS targets
- Verifying cross-platform code paths
- Testing documents and caches directory support

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

Root Directory:
├── iOS-Example.xlsx         # From ios-test (iOS compatibility)
└── [Other-Test].xlsx        # From other platform-specific tests
```

### CI/CD Integration

XLKit includes comprehensive CI/CD testing through GitHub Actions:
- Build & Test: Automated testing on macOS with all unit tests
- Security Scanning: CodeQL analysis for vulnerability detection
- iOS Compatibility: Dedicated iOS testing workflow (`cli-ios.yml`)
- Image Embedding: Automated image embedding tests with validation
- Cross-Platform: macOS and iOS target compilation and testing

### Test Coverage

- Unit Tests: All public APIs and core functionality
- Integration Tests: Complete Excel file generation workflows
- Image Embedding Tests: Aspect ratio preservation for all 17 supported formats
- CSV/TSV Tests: Import/export functionality with various data types
- Performance Tests: Large dataset handling and memory management
- Validation Tests: CoreXLSX compliance verification for all generated files
- Security Tests: Rate limiting, input validation, file quarantine, and checksum verification
- Platform Tests: iOS compatibility and sandbox restrictions testing
- All text alignment options (horizontal, vertical, combined) are fully tested

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