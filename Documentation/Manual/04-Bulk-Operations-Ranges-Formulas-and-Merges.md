# Chapter 04 — Bulk Operations, Ranges, Formulas, and Merges

Navigation: [← Chapter 03](03-Core-Model-Workbook-Sheet-and-Cells.md) · [Manual index](README.md) · [Chapter 05 →](05-CSV-and-TSV.md)

## Basic Usage

### API Highlights

- Workbook: `Workbook()`, `addSheet(name:)`, `save(to:)`, image management methods
- Sheet: `setCell`, `setRow`, `setColumn`, `setRange`, `mergeCells`, `embedImageAutoSized`, `setColumnWidth`
- Convenience Methods: Type-specific setters like `setCell(string:format:)`, `setRange(number:format:)`
- Images: GIF, PNG, JPEG with perfect aspect ratio preservation
- CSV/TSV: `Workbook(fromCSV:)`, `Workbook(fromTSV:)`, `exportToCSV()`, `exportToTSV()`, `importCSV(_:into:hasHeader:)`, `importTSV(_:into:hasHeader:)`
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


### Multiple sheets with formulas

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

### Working with ranges

```swift
let sheet = workbook.addSheet(name: "Range Test")

// Set values in ranges with typed convenience methods (recommended)
sheet.setRange("A1:C3", string: "Range", format: CellFormat.bordered())
sheet.setRange("D1:F3", number: 42.5, format: CellFormat.currency())
sheet.setRange("G1:I3", integer: 100)
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

// Complex border and merge combination
var borderedFormat = CellFormat.bordered()
borderedFormat.fontSize = 11
borderedFormat.fontWeight = .bold
borderedFormat.horizontalAlignment = .center
borderedFormat.verticalAlignment = .center
borderedFormat.fontName = "Calibri"
borderedFormat.borderTop = .thin
borderedFormat.borderBottom = .thin
borderedFormat.borderLeft = .thin
borderedFormat.borderRight = .thin
borderedFormat.borderColor = "#000000"

sheet.setCell("A1", string: "Test1", format: borderedFormat)
sheet.setCell("A2", string: "Test2")
sheet.mergeCells("A1:B1")
sheet.mergeCells("A2:B2")
```
