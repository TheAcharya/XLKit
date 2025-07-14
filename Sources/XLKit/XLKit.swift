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
/// // Create workbook
/// let workbook = Workbook()
/// 
/// // Create workbook from CSV
/// let workbook = Workbook(fromCSV: csvData)
/// 
/// // Add sheet and data
/// let sheet = workbook.addSheet(name: "Data")
///     .setRow(1, strings: ["Name", "Age", "Salary"])
///     .setRow(2, strings: ["Alice", "30", "50000"])
/// 
/// // Save workbook
/// try workbook.save(to: url)
/// 
/// // Export to CSV
/// let csv = sheet.exportToCSV()
/// 
/// // Embed image with perfect aspect ratio preservation
/// try await sheet.embedImageAutoSized(imageData, at: "B2", of: workbook)
/// ```
/// 
/// ## Features
/// - **Excel (.xlsx) file generation** with full OpenXML compliance
/// - **Image embedding** with perfect aspect ratio preservation
/// - **CSV/TSV import/export** with automatic data type detection
/// - **Cell formatting** with comprehensive styling options
/// - **Security features** including rate limiting and file validation
/// - **Cross-platform** support for macOS and iOS
/// - **Modern Swift API** with fluent method chaining
public struct XLKit {
    // This struct is now just a namespace for re-exports
    // All functionality is available as instance methods on Workbook and Sheet
}
