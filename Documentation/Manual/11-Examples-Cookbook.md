# Chapter 11 — Examples Cookbook

Navigation: [← Chapter 10](10-Testing-Test-Runner-CI-and-Code-Style.md) · [Manual index](README.md) · [Chapter 12 →](12-Complete-API-Reference.md)

## Examples

Complete, runnable examples you can copy and adapt.

### Example 1: Minimal workbook and save

Create a workbook, add a sheet, set one cell, and save (sync or async).

```swift
import XLKit

let workbook = Workbook()
let sheet = workbook.addSheet(name: "Report")
sheet.setCell("A1", string: "Hello, XLKit!")

// Synchronous save
let url = CoreUtils.safeFileURL(for: "report.xlsx")
try workbook.save(to: url)

// Or asynchronous
// try await workbook.save(to: url)
```

### Example 2: CSV round-trip

Build a workbook from a CSV string, then export a sheet back to CSV.

```swift
import XLKit

let csv = """
Name,Age,City
Alice,30,London
Bob,25,Paris
"""
let workbook = Workbook(fromCSV: csv, sheetName: "People", hasHeader: true)
guard let sheet = workbook.getSheet(name: "People") else { return }

// Export back to CSV
let exported = sheet.exportToCSV()
print(exported)

// Or export to TSV
let tsv = sheet.exportToTSV()
```

### Example 3: Image embed and save

Load an image from disk, embed it in a cell with automatic sizing, and save the workbook.

```swift
import XLKit

let workbook = Workbook()
let sheet = workbook.addSheet(name: "Photos")
sheet.setRow(1, strings: ["ID", "Photo"])
sheet.setCell("A1", string: "ID", format: CellFormat.header())
sheet.setCell("B1", string: "Photo", format: CellFormat.header())
sheet.setCell("A2", integer: 1)

let imageURL = URL(fileURLWithPath: "/path/to/photo.png")
let imageData = try Data(contentsOf: imageURL)
try await sheet.embedImageAutoSized(imageData, at: "B2", of: workbook)

let outURL = CoreUtils.safeFileURL(for: "photos.xlsx")
try await workbook.save(to: outURL)
```

### Example 4: Multi-sheet with formulas

Two sheets: a data sheet and a summary sheet that references it with formulas.

```swift
import XLKit

let workbook = Workbook()
let dataSheet = workbook.addSheet(name: "Data")
let summarySheet = workbook.addSheet(name: "Summary")

// Data sheet
dataSheet.setCell("A1", string: "Item", format: CellFormat.header())
dataSheet.setCell("B1", string: "Price", format: CellFormat.header())
dataSheet.setCell("A2", string: "Apple")
dataSheet.setCell("B2", number: 1.99, format: CellFormat.currency())
dataSheet.setCell("A3", string: "Orange")
dataSheet.setCell("B3", number: 2.49, format: CellFormat.currency())

// Summary sheet with formulas
summarySheet.setCell("A1", string: "Total", format: CellFormat.header())
summarySheet.setCell("B1", formula: "=SUM(Data!B2:B3)", format: CellFormat.currency())
summarySheet.setCell("A2", string: "Average", format: CellFormat.header())
summarySheet.setCell("B2", formula: "=AVERAGE(Data!B2:B3)", format: CellFormat.currency())

try workbook.save(to: CoreUtils.safeFileURL(for: "summary.xlsx"))
```

### Example 5: Error handling

Catch and handle XLKit errors when saving or validating.

```swift
import XLKit

do {
    let workbook = Workbook()
    let sheet = workbook.addSheet(name: "Sheet1")
    sheet.setCell("A1", string: "Test")
    try workbook.save(to: someURL)
} catch XLKitError.invalidCoordinate(let coord) {
    print("Invalid coordinate: \(coord)")
} catch XLKitError.fileWriteError(let message) {
    print("File write error: \(message)")
} catch XLKitError.zipCreationError(let message) {
    print("ZIP error: \(message)")
} catch XLKitError.securityError(let message) {
    print("Security error: \(message)")
} catch XLKitError.rateLimitExceeded(let message) {
    print("Rate limit: \(message)")
} catch {
    print("Unexpected error: \(error)")
}
```

### Example 6: Hidden and protected sheets

Hide a technical sheet from the tab bar and protect a data sheet so cells are read-only in Excel.

```swift
import XLKit

let workbook = Workbook()
let strings = workbook.addSheet(name: "Strings")
let data = workbook.addSheet(name: "Data")

// Hide the auxiliary sheet (user can unhide in Excel)
strings.state = .hidden
strings.setCell("A1", string: "Internal lookup table")

// Protect the data sheet (locked cells become read-only)
data.setCell("A1", string: "Revenue", format: CellFormat.header())
data.setCell("B1", number: 125_000, format: CellFormat.currency())
data.protection = SheetProtection()

try workbook.save(to: CoreUtils.safeFileURL(for: "protected-report.xlsx"))
```

Use `.veryHidden` when the sheet must not be unhidden from the Excel UI. Set individual `SheetProtection` flags to `false` to explicitly allow specific actions on a protected sheet (see [Chapter 03](03-Core-Model-Workbook-Sheet-and-Cells.md)).

