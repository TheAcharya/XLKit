//
//  CSVUtils.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
@preconcurrency import XLKitCore
import TextFileTools

/// CSV/TSV import/export utilities for XLKit
/// Uses swift-textfile-tools library for robust CSV/TSV parsing and generation
public struct CSVUtils {
    
    /// Exports a sheet to CSV format
    public static func exportToCSV(sheet: Sheet, separator: String = ",") -> String {
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
        
        // Use TextFileTools for CSV generation
        if separator == "," {
            let csv = TextFile.CSV(table: stringTable)
            return csv.rawText
        } else {
            // For custom separators, we need to use a delimited format
            // Since TextFileTools only supports CSV (comma) and TSV (tab),
            // we'll fall back to manual generation for custom separators
            return generateCustomDelimitedText(table: stringTable, separator: separator)
        }
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
        
        // Use TextFileTools for TSV generation
        let tsv = TextFile.TSV(table: stringTable)
        return tsv.rawText
    }
    
    /// Imports CSV data into a sheet
    public static func importFromCSV(sheet: Sheet, csvData: String, separator: String = ",", hasHeader: Bool = false) {
        // Use TextFileTools for CSV parsing
        let stringTable: StringTable
        if separator == "," {
            let csv = TextFile.CSV(rawText: csvData)
            stringTable = csv.table
        } else {
            // For custom separators, parse manually
            stringTable = parseCustomDelimitedText(text: csvData, separator: separator)
        }
        
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
        // Use TextFileTools for TSV parsing
        let tsv = TextFile.TSV(rawText: tsvData)
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
    
    /// Creates a workbook from CSV data
    public static func createWorkbookFromCSV(csvData: String, sheetName: String = "Sheet1", separator: String = ",", hasHeader: Bool = false) -> Workbook {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: sheetName)
        importFromCSV(sheet: sheet, csvData: csvData, separator: separator, hasHeader: hasHeader)
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
    
    /// Generates custom delimited text for separators other than comma or tab
    /// Falls back to manual generation when TextFileTools doesn't support the separator
    private static func generateCustomDelimitedText(table: StringTable, separator: String) -> String {
        return table.map { row in
            row.map { textString in
                var outString = textString
                
                // Escape double-quotes
                outString = outString.replacingOccurrences(of: "\"", with: "\"\"")
                
                // Wrap string in double-quotes if it contains separator, quotes, or newlines
                if outString.contains(separator) ||
                   outString.contains("\"") ||
                   outString.contains("\n") ||
                   outString.contains("\r") {
                    outString = "\"\(outString)\""
                }
                
                return outString
            }
            .joined(separator: separator)
        }
        .joined(separator: "\n")
    }
    
    /// Parses custom delimited text for separators other than comma or tab
    /// Falls back to manual parsing when TextFileTools doesn't support the separator
    private static func parseCustomDelimitedText(text: String, separator: String) -> StringTable {
        let lines = text.components(separatedBy: .newlines)
            .filter { !$0.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty }
        
        var result: StringTable = []
        
        for line in lines {
            var fields: [String] = []
            var currentField = ""
            var inQuotes = false
            var i = 0
            
            while i < line.count {
                let char = line[line.index(line.startIndex, offsetBy: i)]
                
                if char == "\"" {
                    if inQuotes {
                        // Check for escaped quote
                        if i + 1 < line.count && line[line.index(line.startIndex, offsetBy: i + 1)] == "\"" {
                            currentField += "\""
                            i += 1 // Skip the next quote
                        } else {
                            inQuotes = false
                        }
                    } else {
                        inQuotes = true
                    }
                } else if String(char) == separator && !inQuotes {
                    fields.append(currentField)
                    currentField = ""
                } else {
                    currentField += String(char)
                }
                
                i += 1
            }
            
            fields.append(currentField)
            result.append(fields)
        }
        
        return result
    }
}

// MARK: - String Extension for Pattern Matching

private extension String {
    func matches(pattern: String) -> Bool {
        return range(of: pattern, options: .regularExpression) != nil
    }
} 