//
//  CSVUtils.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
@preconcurrency import XLKitCore
import TextFile

/// CSV/TSV import/export utilities for XLKit.
///
/// Only CSV (comma) and TSV (tab) are supported. These are defined formats with
/// standard quote and escape rules (e.g. RFC 4180 for CSV). Custom delimiters
/// are not supported to avoid non-standard formats and delimiter-vs-content
/// collision pitfalls.
///
/// Uses the swift-textfile library for spec-compliant CSV/TSV parsing and generation.
public struct CSVUtils {

    /// Exports a sheet to CSV format (comma-separated, RFC 4180-style).
    public static func exportToCSV(sheet: Sheet) -> String {
        // Find max row/column from used cells
        let usedCells = sheet.getUsedCells()
        var maxRow = 0
        var maxColumn = 0

        for coordinate in usedCells {
            guard let cellCoord = CellCoordinate(excelAddress: coordinate) else { continue }
            maxRow = max(maxRow, cellCoord.row)
            maxColumn = max(maxColumn, cellCoord.column)
        }

        // Build StringTable from sheet data
        var stringTable: StringTable = []

        for row in 1...maxRow {
            var rowData: [String] = []

            for column in 1...maxColumn {
                let coord = CellCoordinate(row: row, column: column)
                let cellAddress = coord.excelAddress

                if let cellValue = sheet.getCell(cellAddress) {
                    rowData.append(cellValue.stringValue)
                } else {
                    rowData.append("")
                }
            }

            stringTable.append(rowData)
        }

        let csv = CSV(table: stringTable)
        return csv.rawText
    }
    
    /// Exports a sheet to TSV format
    public static func exportToTSV(sheet: Sheet) -> String {
        // Find max row/column from used cells
        let usedCells = sheet.getUsedCells()
        var maxRow = 0
        var maxColumn = 0
        
        for coordinate in usedCells {
            guard let cellCoord = CellCoordinate(excelAddress: coordinate) else { continue }
            maxRow = max(maxRow, cellCoord.row)
            maxColumn = max(maxColumn, cellCoord.column)
        }
        
        // Build StringTable from sheet data
        var stringTable: StringTable = []
        
        for row in 1...maxRow {
            var rowData: [String] = []
            
            for column in 1...maxColumn {
                let coord = CellCoordinate(row: row, column: column)
                let cellAddress = coord.excelAddress
                
                if let cellValue = sheet.getCell(cellAddress) {
                    rowData.append(cellValue.stringValue)
                } else {
                    rowData.append("")
                }
            }
            
            stringTable.append(rowData)
        }
        
        // Use TextFile for TSV generation
        let tsv = TSV(table: stringTable)
        return tsv.rawText
    }
    
    /// Imports CSV data into a sheet (comma-separated, RFC 4180-style).
    public static func importFromCSV(sheet: Sheet, csvData: String, hasHeader: Bool = false) {
        let csv = CSV(rawText: csvData)
        let stringTable = csv.table

        // Determine which rows to import
        let dataRows: [[String]]
        if hasHeader && !stringTable.isEmpty {
            dataRows = Array(stringTable.dropFirst())
        } else {
            dataRows = stringTable
        }
        let startRow = hasHeader ? 2 : 1
        
        // Import data into sheet
        for (rowIndex, row) in dataRows.enumerated() {
            let excelRow = startRow + rowIndex
            
            for (columnIndex, value) in row.enumerated() {
                let column = columnIndex + 1
                let coord = CellCoordinate(row: excelRow, column: column)
                let cellAddress = coord.excelAddress
                
                let cellValue = parseCSVValue(value)
                sheet.setCell(cellAddress, value: cellValue)
            }
        }
    }
    
    /// Imports TSV data into a sheet
    public static func importFromTSV(sheet: Sheet, tsvData: String, hasHeader: Bool = false) {
        // Use TextFile for TSV parsing
        let tsv = TSV(rawText: tsvData)
        let stringTable = tsv.table
        
        // Determine which rows to import
        let dataRows: [[String]]
        if hasHeader && !stringTable.isEmpty {
            dataRows = Array(stringTable.dropFirst())
        } else {
            dataRows = stringTable
        }
        let startRow = hasHeader ? 2 : 1
        
        // Import data into sheet
        for (rowIndex, row) in dataRows.enumerated() {
            let excelRow = startRow + rowIndex
            
            for (columnIndex, value) in row.enumerated() {
                let column = columnIndex + 1
                let coord = CellCoordinate(row: excelRow, column: column)
                let cellAddress = coord.excelAddress
                
                let cellValue = parseCSVValue(value)
                sheet.setCell(cellAddress, value: cellValue)
            }
        }
    }
    
    /// Creates a workbook from CSV data (comma-separated).
    public static func createWorkbookFromCSV(csvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) -> Workbook {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: sheetName)
        importFromCSV(sheet: sheet, csvData: csvData, hasHeader: hasHeader)
        return workbook
    }
    
    /// Creates a workbook from TSV data
    public static func createWorkbookFromTSV(tsvData: String, sheetName: String = "Sheet1", hasHeader: Bool = false) -> Workbook {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: sheetName)
        importFromTSV(sheet: sheet, tsvData: tsvData, hasHeader: hasHeader)
        return workbook
    }
    
    // MARK: - Private Helper Methods
    
    /// Parses CSV value to appropriate CellValue type
    private static func parseCSVValue(_ value: String) -> CellValue {
        let trimmed = value.trimmingCharacters(in: .whitespaces)
        
        // Check for boolean values
        if trimmed.lowercased() == "true" {
            return .boolean(true)
        } else if trimmed.lowercased() == "false" {
            return .boolean(false)
        }
        
        // Check for integer
        if let intValue = Int(trimmed) {
            return .integer(intValue)
        }
        
        // Check for double
        if let doubleValue = Double(trimmed) {
            return .number(doubleValue)
        }
        
        // Check for date (yyyy-MM-dd format)
        if trimmed.matches(pattern: "^\\d{4}-\\d{2}-\\d{2}$") {
            let formatter = DateFormatter()
            formatter.dateFormat = "yyyy-MM-dd"
            if let date = formatter.date(from: trimmed) {
                return .date(date)
            }
        }
        
        // Default to string
        return .string(trimmed)
    }
    
}

// MARK: - String Extension for Pattern Matching

private extension String {
    func matches(pattern: String) -> Bool {
        return range(of: pattern, options: .regularExpression) != nil
    }
} 
