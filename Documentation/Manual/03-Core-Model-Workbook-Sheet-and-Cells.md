# Chapter 03 — Core Model: Workbook, Sheet, and Cells

Navigation: [← Chapter 02](02-Architecture-Modules-and-Source-Map.md) · [Manual index](README.md) · [Chapter 04 →](04-Bulk-Operations-Ranges-Formulas-and-Merges.md)

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

// Border formatting
var borderedFormat = CellFormat.bordered()
borderedFormat.borderTop = .thin
borderedFormat.borderBottom = .thin
borderedFormat.borderLeft = .thin
borderedFormat.borderRight = .thin
borderedFormat.borderColor = "#000000"
sheet.setCell("A1", string: "Bordered Cell", format: borderedFormat)

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

