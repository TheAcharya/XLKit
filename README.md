# XLKit

A universal, generic Swift library for creating and manipulating Excel (.xlsx) files on macOS.

## Features

- **Universal & Generic**: No font or icon dependencies - pure data manipulation
- **Swift 6.0 Ready**: Modern Swift with async/await support
- **Native ZIP**: Uses Foundation's Archive for ZIP creation
- **Excel-Compatible**: Generates standard .xlsx files
- **Type Safe**: Strong typing with enums and structs
- **Memory Efficient**: Optimized for large datasets

## Requirements

- macOS 13.0+
- Swift 6.0+
- Xcode 15.0+

## Installation

### Swift Package Manager

Add XLKit to your project dependencies:

```swift
dependencies: [
    .package(url: "https://github.com/yourusername/XLKit.git", from: "1.0.0")
]
```

## Quick Start

```swift
import XLKit

// Create a new workbook
let workbook = XLKit.createWorkbook()

// Add a sheet
let sheet = workbook.addSheet(name: "Data")

// Add data
sheet.setCell("A1", value: .string("Name"))
sheet.setCell("B1", value: .string("Age"))
sheet.setCell("C1", value: .string("Salary"))

sheet.setCell("A2", value: .string("John"))
sheet.setCell("B2", value: .integer(30))
sheet.setCell("C2", value: .number(50000.0))

// Save the workbook
try await XLKit.saveWorkbook(workbook, to: fileURL)
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
- Microsoft Excel
- Google Sheets
- LibreOffice Calc
- Numbers (macOS)
- Any application that supports the OpenXML format

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any issues or have questions, please open an issue on GitHub.
