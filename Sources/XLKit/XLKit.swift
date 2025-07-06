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

import Foundation

public struct XLKit {
    /// Creates a new workbook
    public static func createWorkbook() -> Workbook {
        return Workbook()
    }
    /// Saves a workbook to a file asynchronously
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) async throws {
        try await XLSXEngine.generateXLSX(workbook: workbook, to: url)
    }
    /// Saves a workbook to a file synchronously
    public static func saveWorkbook(_ workbook: Workbook, to url: URL) throws {
        try XLSXEngine.generateXLSXSync(workbook: workbook, to: url)
    }
    // MARK: - CSV/TSV Convenience Methods
    /// Creates a workbook from CSV data
    public static func createWorkbookFromCSV(csvData: String, sheetName: String = "Sheet1", separator: String = ",", hasHeader: Bool = false) -> Workbook {
        return CSVUtils.createWorkbookFromCSV(csvData: csvData, sheetName: sheetName, separator: separator, hasHeader: hasHeader)
    }
    /// Creates a workbook from TSV data
    public static func createWorkbookFromTSV(tsvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) -> Workbook {
        return CSVUtils.createWorkbookFromTSV(tsvData: tsvData, sheetName: sheetName, hasHeader: hasHeader)
    }
    /// Exports a sheet to CSV format
    public static func exportSheetToCSV(sheet: Sheet, separator: String = ",") -> String {
        return CSVUtils.exportToCSV(sheet: sheet, separator: separator)
    }
    /// Exports a sheet to TSV format
    public static func exportSheetToTSV(sheet: Sheet) -> String {
        return CSVUtils.exportToTSV(sheet: sheet)
    }
    /// Imports CSV data into a sheet
    public static func importCSVIntoSheet(sheet: Sheet, csvData: String, separator: String = ",", hasHeader: Bool = false) {
        CSVUtils.importFromCSV(sheet: sheet, csvData: csvData, separator: separator, hasHeader: hasHeader)
    }
    /// Imports TSV data into a sheet
    public static func importTSVIntoSheet(sheet: Sheet, tsvData: String, hasHeader: Bool = false) {
        CSVUtils.importFromTSV(sheet: sheet, tsvData: tsvData, hasHeader: hasHeader)
    }
}
