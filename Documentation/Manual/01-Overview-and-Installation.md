# Chapter 01 — Overview & Installation

Navigation: [Manual index](README.md) · Next: [Chapter 02 — Architecture](02-Architecture-Modules-and-Source-Map.md)

XLKit is a Swift library for creating Excel `.xlsx` workbooks on **macOS 12+** and **iOS 15+**. It provides a fluent API for cells, formatting, CSV/TSV, images, sheet visibility, sheet protection, and OpenXML-compliant export. This chapter covers performance, compatibility, requirements, and installation.

## Performance Considerations

- Memory Usage: XLKit is optimized for large datasets with efficient memory management
- Async Operations: Use async/await for file operations to avoid blocking the main thread
- Batch Operations: Set multiple cells in batches for better performance
- Range Operations: Use `setRange()` for setting multiple cells with the same value

## File format

XLKit generates **Excel `.xlsx`** workbooks (Office Open XML). It does **not** export Apple **Numbers** `.numbers` files or any format other than `.xlsx`.

Files produced are intended for apps that read standard Excel workbooks, for example **Microsoft Excel**, **LibreOffice Calc**, or **Google Sheets** (via upload/import). Compatibility and rendering can vary by app; validate in your target environment if you rely on a specific viewer.

## Requirements

- macOS: 12.0+
- iOS: 15.0+ (available but not tested)
- Swift: 6.0+


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
    .package(url: "https://github.com/TheAcharya/XLKit.git", from: "1.1.6")
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
sheet.setCell("C1", string: "Age", format: CellFormat.text(color: "#FF0000"))

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

