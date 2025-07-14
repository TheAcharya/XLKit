//
//  XLKit.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

@_exported import XLKitCore
@_exported import XLKitFormatters
@_exported import XLKitImages
@_exported import XLKitXLSX

/// XLKit main API: re-exports all public APIs from submodules for easy use.
/// 
/// XLKit provides a modern Swift API for creating and manipulating Excel (.xlsx) files
/// with advanced features including image embedding, CSV/TSV import/export, and security.
/// 
/// ## Quick Start
/// 
/// ```swift
/// import XLKit
/// 
/// // Create workbook
/// let workbook = Workbook()
/// 
/// // Create workbook from CSV/TSV
/// let workbook = Workbook(fromCSV: csvData, hasHeader: true)
/// let workbook = Workbook(fromTSV: tsvData, hasHeader: true)
/// 
/// // Add sheet and data with convenience methods
/// let sheet = workbook.addSheet(name: "Data")
///     .setCell("A1", string: "Name", format: CellFormat.header())
///     .setCell("B1", string: "Age", format: CellFormat.header())
///     .setCell("C1", string: "Salary", format: CellFormat.header())
///     .setRow(2, strings: ["Alice", "30", "50000"])
///     .setRow(3, strings: ["Bob", "25", "45000"])
/// 
/// // Add formatted data
/// sheet.setCell("A2", string: "Alice")
/// sheet.setCell("B2", integer: 30, format: CellFormat.number())
/// sheet.setCell("C2", number: 50000.0, format: CellFormat.currency())
/// 
/// // Embed image with perfect aspect ratio preservation
/// let imageData = try Data(contentsOf: URL(fileURLWithPath: "photo.png"))
/// try await sheet.embedImageAutoSized(imageData, at: "D2", of: workbook)
/// 
/// // Save workbook (sync or async)
/// try await workbook.save(to: url)
/// // or
/// try workbook.save(to: url)
/// 
/// // Export to CSV/TSV
/// let csv = sheet.exportToCSV()
/// let tsv = sheet.exportToTSV()
/// ```
/// 
/// ## Core Features
/// 
/// ### Workbook Management
/// ```swift
/// let workbook = Workbook()
/// let sheet = workbook.addSheet(name: "Sheet1")
/// let allSheets = workbook.getSheets()
/// let specificSheet = workbook.getSheet(name: "Sheet1")
/// workbook.removeSheet(name: "Sheet1")
/// ```
/// 
/// ### Cell Operations
/// ```swift
/// // Basic cell setting
/// sheet.setCell("A1", value: .string("Hello"))
/// sheet.setCell("B1", value: .number(42.5))
/// 
/// // Convenience methods with formatting (recommended)
/// sheet.setCell("A1", string: "Header", format: CellFormat.header())
/// sheet.setCell("B1", number: 1234.56, format: CellFormat.currency())
/// sheet.setCell("C1", integer: 100, format: CellFormat.number())
/// sheet.setCell("D1", boolean: true, format: CellFormat.text())
/// sheet.setCell("E1", date: Date(), format: CellFormat.date())
/// sheet.setCell("F1", formula: "=A1+B1", format: CellFormat.text())
/// 
/// // Get cell values
/// let value = sheet.getCell("A1")
/// let cellWithFormat = sheet.getCellWithFormat("A1")
/// ```
/// 
/// ### Range Operations
/// ```swift
/// // Set ranges with convenience methods
/// sheet.setRange("A1:C3", string: "Range", format: CellFormat.bordered())
/// sheet.setRange("D1:F3", number: 42.5, format: CellFormat.currency())
/// sheet.setRange("G1:I3", integer: 100, format: CellFormat.number())
/// 
/// // Merge cells
/// sheet.mergeCells("A1:C1")
/// let mergedRanges = sheet.getMergedRanges()
/// ```
/// 
/// ### Image Embedding
/// ```swift
/// // Automatic sizing with perfect aspect ratio preservation
/// try await sheet.embedImageAutoSized(imageData, at: "A1", of: workbook)
/// 
/// // With scaling control
/// try await sheet.embedImage(imageData, at: "B1", of: workbook, scale: 0.7)
/// 
/// // From file
/// try await sheet.embedImage(from: imageURL, at: "C1")
/// 
/// // Manual image management
/// let excelImage = ExcelImage(id: "img1", data: imageData, format: .png, originalSize: size)
/// sheet.addImage(excelImage, at: "D1")
/// workbook.addImage(excelImage)
/// ```
/// 
/// ### CSV/TSV Operations
/// ```swift
/// // Import data
/// workbook.importCSV(csvData, into: sheet, hasHeader: true)
/// workbook.importTSV(tsvData, into: sheet, hasHeader: true)
/// 
/// // Export data
/// let csv = sheet.exportToCSV()
/// let tsv = sheet.exportToTSV()
/// let workbookCSV = workbook.exportSheetToCSV(sheet)
/// let workbookTSV = workbook.exportSheetToTSV(sheet)
/// ```
/// 
/// ### Utility Properties
/// ```swift
/// // Sheet utilities
/// let allCells = sheet.allCells                    // [String: CellValue]
/// let allFormattedCells = sheet.allFormattedCells  // [String: Cell]
/// let isEmpty = sheet.isEmpty                      // Bool
/// let cellCount = sheet.cellCount                  // Int
/// let imageCount = sheet.imageCount                // Int
/// 
/// // Workbook utilities
/// let workbookImageCount = workbook.imageCount     // Int
/// let allImages = workbook.getImages()             // [ExcelImage]
/// ```
/// 
/// ## Features
/// - **Excel (.xlsx) file generation** with full OpenXML compliance
/// - **Image embedding** with perfect aspect ratio preservation for all 17 professional video/cinema formats
/// - **CSV/TSV import/export** with automatic data type detection
/// - **Cell formatting** with comprehensive styling options and predefined formats
/// - **Security features** including rate limiting, file validation, and checksums
/// - **Cross-platform** support for macOS and iOS
/// - **Modern Swift API** with fluent method chaining and type-safe convenience methods
/// - **Comprehensive testing** with automated validation using CoreXLSX
public struct XLKit {
    // This struct serves as a namespace for re-exports
    // All functionality is available as instance methods on Workbook and Sheet classes
}
