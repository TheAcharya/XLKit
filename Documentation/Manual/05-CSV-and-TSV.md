# Chapter 05 — CSV and TSV

Navigation: [← Chapter 04](04-Bulk-Operations-Ranges-Formulas-and-Merges.md) · [Manual index](README.md) · [Chapter 06 →](06-Images-Embedding-and-Sizing.md)

## CSV/TSV Import & Export

XLKit provides simple static methods for importing and exporting CSV/TSV data. Only CSV (comma) and TSV (tab) are supported; both are defined formats with standard quote and escape rules (e.g. RFC 4180 for CSV). Parsing and generation are powered by the [swift-textfile](https://github.com/orchetect/swift-textfile) library for spec-compliant handling of quoted fields and escaped quotes.

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

// Import CSV into an existing sheet (Workbook method)
workbook.importCSV(csvData, into: sheet, hasHeader: true)

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

// Import TSV into an existing sheet (Workbook method)
tsvWorkbook.importTSV(tsvData, into: tsvSheet, hasHeader: true)
```

All CSV/TSV helpers are available as instance methods on `Workbook` and `Sheet` classes for convenience, and are powered by the `XLKitFormatters` module under the hood.

