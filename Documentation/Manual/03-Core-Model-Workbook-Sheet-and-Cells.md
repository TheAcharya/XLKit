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

// Sheet visibility (tab bar)
sheet.state = .visible      // default — shown in tab bar
sheet.state = .hidden         // hidden; user can unhide from Excel UI
sheet.state = .veryHidden     // hidden; cannot be unhidden from Excel UI

// Sheet protection (Excel "Protect Sheet")
sheet.protection = SheetProtection()  // enable protection with XLSX defaults
var protection = SheetProtection()
protection.password = "CC3F"          // optional legacy password hash
protection.selectLockedCells = true   // true = lock that action (inverted semantics)
sheet.protection = protection

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

### Sheet visibility and protection

Sheets can be marked as auxiliary (hidden from the tab bar) or read-only (protected) when the workbook opens in Excel, LibreOffice Calc, or Google Sheets. Apple Numbers ignores these XLSX features.

#### `SheetState`

| Value | Behaviour |
|-------|-----------|
| `.visible` (default) | Sheet appears in the tab bar. No extra XML is emitted. |
| `.hidden` | Sheet is hidden; the user can unhide it from the workbook UI. |
| `.veryHidden` | Sheet is hidden and cannot be unhidden from the UI (programmatic access only). |

When the first sheet(s) in the workbook are hidden, XLKit emits `activeTab` on `<workbookView>` pointing at the first visible sheet so Excel opens the file correctly.

```swift
let workbook = Workbook()
let tech = workbook.addSheet(name: "Strings")
let data = workbook.addSheet(name: "Data")

tech.state = .hidden
data.setCell("A1", string: "User-facing data")

try workbook.save(to: url)
```

#### `SheetProtection`

Set `sheet.protection` to a non-nil `SheetProtection` to enable Excel's "Protect Sheet" behaviour. Cells with default locked styling become read-only. Set `protection` to `nil` (default) for an unprotected sheet — no `<sheetProtection>` element is emitted.

**Important:** Boolean permission flags use **inverted lock semantics** — `true` means the action is **locked** (not allowed) when the sheet is protected, not "allowed." Omit a flag (`nil`) to use the XLSX-defined default.

```swift
var protection = SheetProtection()
protection.sheet = true              // enforce protection (default)
protection.formatCells = false       // explicitly allow formatting
protection.selectLockedCells = true  // block selection of locked cells
sheet.protection = protection
```

For password-protected sheets, use `CoreUtils` — do not guess salt/hash values:

```swift
var protection = SheetProtection()
try CoreUtils.configureSheetPassword(&protection, plaintext: "mySecret")
sheet.protection = protection
```

- **Legacy only:** `protection.password = CoreUtils.excelLegacySheetPasswordHash(for: "mySecret")`
- **Modern (SHA-512):** `let modern = try CoreUtils.excelModernSheetPasswordHash(for: "mySecret")` then assign `algorithmName`, `saltValue`, `hashValue`, `spinCount`
- **Developer helper:** `swift run XLKitTestRunner sheet-password mySecret` (add `--demo-salts` with **`1234`** to match **`Comprehensive-Demo.xlsx`** — salts are in `ComprehensiveDemoProtection.swift`, TestRunner only)

**Note:** `sheet.clear()` clears cells, formats, images, merges, and dimensions but does **not** reset `state` or `protection`.

