# XLKit

A modern, ultra-easy Swift library for creating and manipulating Excel (.xlsx) files on macOS.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Core Concepts](#core-concepts)
- [Basic Usage](#basic-usage)
- [CSV/TSV Import & Export](#csvtsv-import--export)
- [Image Support](#image-support)
- [Advanced Usage](#advanced-usage)
- [Error Handling](#error-handling)
- [Performance Considerations](#performance-considerations)
- [Code Style & Formatting](#code-style--formatting)
- [Testing](#testing)
- [Contributing](#contributing)
- [Support](#support)
- [License](#license)

## Features

- **Effortless API**: Fluent, chainable, and bulk data helpers
- **Image Embedding**: Add GIF, PNG, JPEG, BMP, TIFF images to any cell
- **Auto Column Sizing**: Set or auto-size columns based on image dimensions
- **Async & Sync Save**: Save workbooks with one line (async or sync)
- **Type-Safe**: Strong enums and structs for all data
- **No Dependencies**: Pure Swift, macOS 13+, Swift 6.0+
- **Comprehensive Tests**: 100% tested, CI ready

## Requirements

- **macOS**: 13.0+
- **Swift**: 6.0+

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

// 3. Add a GIF image to a cell (one-liner)
let gifData = try Data(contentsOf: URL(fileURLWithPath: "alice.gif"))
sheet.addImage(gifData, at: "B2", format: .gif)
    .autoSizeColumn(2, forImageAt: "B2")

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

- **Workbook**: `createWorkbook()`, `addSheet(name:)`, `save(to:)`
- **Sheet**: `setCell`, `setRow`, `setColumn`, `setRange`, `mergeCells`, `addImage`, `autoSizeColumn`, `setColumnWidth`
- **Images**: GIF, PNG, JPEG, BMP, TIFF; add from Data or file URL
- **Fluent API**: Most setters return `Self` for chaining
- **Bulk Data**: `setRow`, `setColumn` for easy import
- **Doc Comments**: All public APIs are documented for Xcode autocomplete

### Example: Bulk Data and Images

```swift
let sheet = workbook.addSheet(name: "Products")
    .setRow(1, values: [.string("Product"), .string("Image"), .string("Price")])
    .setRow(2, values: [.string("Apple"), .empty, .number(1.99)])
    .setRow(3, values: [.string("Banana"), .empty, .number(0.99)])

let appleGif = try Data(contentsOf: URL(fileURLWithPath: "apple.gif"))
sheet.addImage(appleGif, at: "B2", format: .gif)
    .autoSizeColumn(2, forImageAt: "B2")
```

### Column Sizing

```swift
sheet.setColumnWidth(2, width: 200) // Set manually
sheet.autoSizeColumn(2, forImageAt: "B2") // Auto-size to fit image
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

### Supported Image Formats
- **GIF** (including animated)
- **PNG**
- **JPEG/JPG**
- **BMP**
- **TIFF**

### Adding Images

```swift
// Add image from Data
let imageData = try Data(contentsOf: URL(fileURLWithPath: "image.png"))
sheet.addImage(imageData, at: "A1", format: .png)

// Add image from URL
let imageURL = URL(fileURLWithPath: "image.gif")
sheet.addImage(imageURL, at: "B1")

// Auto-size column to fit image
sheet.autoSizeColumn(1, forImageAt: "A1")
```

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

### Date Handling

```swift
let sheet = workbook.addSheet(name: "Dates")

// Add dates
let today = Date()
sheet.setCell("A1", value: .date(today))

// Convert between Excel dates and Swift dates
let excelNumber = XLKitUtils.excelNumberFromDate(today)
let convertedDate = XLKitUtils.dateFromExcelNumber(excelNumber)

// Format dates for display
let formattedDate = XLKitUtils.formatDate(today)
```

### Column/Row Utilities

```swift
// Convert column numbers to letters
XLKitUtils.columnLetter(from: 1)  // "A"
XLKitUtils.columnLetter(from: 26) // "Z"
XLKitUtils.columnLetter(from: 27) // "AA"

// Convert letters to numbers
XLKitUtils.columnNumber(from: "A")  // 1
XLKitUtils.columnNumber(from: "Z")  // 26
XLKitUtils.columnNumber(from: "AA") // 27
```

## Error Handling

XLKit provides comprehensive error handling:

```swift
do {
    try await XLKit.saveWorkbook(workbook, to: fileURL)
} catch XLKitError.invalidCoordinate(let coord) {
    print("Invalid coordinate: \(coord)")
} catch XLKitError.fileWriteError(let message) {
    print("File write error: \(message)")
} catch XLKitError.zipCreationError(let message) {
    print("ZIP creation error: \(message)")
} catch {
    print("Unknown error: \(error)")
}
```

## Performance Considerations

- **Memory Usage**: XLKit is optimized for large datasets with efficient memory management
- **Async Operations**: Use async/await for file operations to avoid blocking the main thread
- **Batch Operations**: Set multiple cells in batches for better performance
- **Range Operations**: Use `setRange()` for setting multiple cells with the same value

## File Format

XLKit generates standard Excel (.xlsx) files that are compatible with:
- **Microsoft Excel**
- **Google Sheets**
- **LibreOffice Calc**
- **Numbers (macOS)**
- Any application that supports the OpenXML format

## Code Style & Formatting

XLKit enforces a modern, consistent Swift code style for all contributions:

- **4-space indentation**
- **Trailing commas**
- **Grouped and reordered imports**
- **120 character line length**
- **Consistent spacing and blank lines**
- **No force-unwraps or force-casts in public API**
- **All public APIs have doc comments**
- **Follows Swift 6 idioms and best practices**

A `.swift-format` file is included in the repo. To format the codebase, run:

```sh
swift-format format -i .
```

Or use Xcode's built-in formatter for most style rules.

All code must be formatted and pass CI before merging. See `.cursorrules` for more details.

## Testing

- **100% tested**, including image and column sizing features
- **Comprehensive test suite** covering all public APIs
- **CI/CD ready** with automated testing

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any issues or have questions, please open an issue on GitHub.

## License

This project is licensed under the MIT License - see the LICENSE file for details.