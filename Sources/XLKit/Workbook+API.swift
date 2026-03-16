//
//  Workbook+API.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKitCore
import XLKitXLSX
import XLKitFormatters

public extension Workbook {
    /// Creates a workbook from CSV data (comma-separated).
    convenience init(fromCSV csvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) {
        self.init()
        let sheet = addSheet(name: sheetName)
        importCSV(csvData, into: sheet, hasHeader: hasHeader)
    }
    /// Creates a workbook from TSV data
    convenience init(fromTSV tsvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) {
        self.init()
        let sheet = addSheet(name: sheetName)
        importTSV(tsvData, into: sheet, hasHeader: hasHeader)
    }
    // MARK: - File Operations
    /// Saves the workbook to a file synchronously
    @MainActor
    func save(to url: URL) throws {
        try XLSXEngine.generateXLSX(workbook: self, to: url)
    }
    /// Saves the workbook to a file asynchronously
    @MainActor
    func save(to url: URL) async throws {
        try XLSXEngine.generateXLSX(workbook: self, to: url)
    }
    // MARK: - CSV/TSV Operations
    /// Imports CSV data into a sheet (comma-separated).
    func importCSV(_ csvData: String, into sheet: Sheet, hasHeader: Bool = false) {
        CSVUtils.importFromCSV(sheet: sheet, csvData: csvData, hasHeader: hasHeader)
    }
    /// Imports TSV data into a sheet
    func importTSV(_ tsvData: String, into sheet: Sheet, hasHeader: Bool = false) {
        CSVUtils.importFromTSV(sheet: sheet, tsvData: tsvData, hasHeader: hasHeader)
    }
    /// Exports a sheet to CSV format (comma-separated).
    func exportSheetToCSV(_ sheet: Sheet) -> String {
        return CSVUtils.exportToCSV(sheet: sheet)
    }
    /// Exports a sheet to TSV format
    func exportSheetToTSV(_ sheet: Sheet) -> String {
        return CSVUtils.exportToTSV(sheet: sheet)
    }
} 