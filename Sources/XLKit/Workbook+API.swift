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
    /// Creates a workbook from CSV data
    convenience init(fromCSV csvData: String, sheetName: String = "Sheet1", separator: String = ",", hasHeader: Bool = false) {
        self.init()
        let sheet = addSheet(name: sheetName)
        importCSV(csvData, into: sheet, separator: separator, hasHeader: hasHeader)
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
    /// Imports CSV data into a sheet
    func importCSV(_ csvData: String, into sheet: Sheet, separator: String = ",", hasHeader: Bool = false) {
        CSVUtils.importFromCSV(sheet: sheet, csvData: csvData, separator: separator, hasHeader: hasHeader)
    }
    /// Imports TSV data into a sheet
    func importTSV(_ tsvData: String, into sheet: Sheet, hasHeader: Bool = false) {
        CSVUtils.importFromTSV(sheet: sheet, tsvData: tsvData, hasHeader: hasHeader)
    }
    /// Exports a sheet to CSV format
    func exportSheetToCSV(_ sheet: Sheet, separator: String = ",") -> String {
        return CSVUtils.exportToCSV(sheet: sheet, separator: separator)
    }
    /// Exports a sheet to TSV format
    func exportSheetToTSV(_ sheet: Sheet) -> String {
        return CSVUtils.exportToTSV(sheet: sheet)
    }
} 