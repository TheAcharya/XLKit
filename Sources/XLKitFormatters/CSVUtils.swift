//
//  CSVUtils.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
@preconcurrency import XLKitCore

// MARK: - CSV/TSV Utilities for XLKit

/// CSV/TSV import/export utilities for XLKit
public struct CSVUtils {
    
    /// Exports a sheet to CSV format
    public static func exportToCSV(sheet: Sheet, separator: String = ",") -> String {
        var csv = ""
        
        // Get all used cells and find max row/column
        let usedCells = sheet.getUsedCells()
        var maxRow = 0
        var maxColumn = 0
        
        for coordinate in usedCells {
            guard let cellCoord = CellCoordinate(excelAddress: coordinate) else { continue }
            maxRow = max(maxRow, cellCoord.row)
            maxColumn = max(maxColumn, cellCoord.column)
        }
        
        // Generate CSV content
        for row in 1...maxRow {
            var rowData: [String] = []
            
            for column in 1...maxColumn {
                let coord = CellCoordinate(row: row, column: column)
                let cellAddress = coord.excelAddress
                
                if let cellValue = sheet.getCell(cellAddress) {
                    rowData.append(formatCellValueForCSV(cellValue, separator: separator))
                } else {
                    rowData.append("")
                }
            }
            
            csv += rowData.joined(separator: separator) + "\n"
        }
        
        return csv.trimmingCharacters(in: .whitespacesAndNewlines)
    }
    
    /// Exports a sheet to TSV format
    public static func exportToTSV(sheet: Sheet) -> String {
        return exportToCSV(sheet: sheet, separator: "\t")
    }
    
    /// Imports CSV data into a sheet
    public static func importFromCSV(sheet: Sheet, csvData: String, separator: String = ",", hasHeader: Bool = false) {
        let lines = csvData.components(separatedBy: .newlines)
            .filter { !$0.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty }
        
        let dataLines: [String]
        if hasHeader && !lines.isEmpty {
            dataLines = Array(lines.dropFirst())
        } else {
            dataLines = lines
        }
        let startRow = hasHeader ? 2 : 1
        
        for (lineIndex, line) in dataLines.enumerated() {
            let row = startRow + lineIndex
            let columns = parseCSVLine(line, separator: separator)
            
            for (columnIndex, value) in columns.enumerated() {
                let column = columnIndex + 1
                let coord = CellCoordinate(row: row, column: column)
                let cellAddress = coord.excelAddress
                
                let cellValue = parseCSVValue(value)
                sheet.setCell(cellAddress, value: cellValue)
            }
        }
    }
    
    /// Imports TSV data into a sheet
    public static func importFromTSV(sheet: Sheet, tsvData: String, hasHeader: Bool = false) {
        importFromCSV(sheet: sheet, csvData: tsvData, separator: "\t", hasHeader: hasHeader)
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
        return createWorkbookFromCSV(csvData: tsvData, sheetName: sheetName, separator: "\t", hasHeader: hasHeader)
    }
    
    // MARK: - Private Helper Methods
    
    private static func formatCellValueForCSV(_ cellValue: CellValue, separator: String) -> String {
        let stringValue = cellValue.stringValue
        
        // Check if we need to quote the value
        let needsQuoting = stringValue.contains(separator) || 
                          stringValue.contains("\"") || 
                          stringValue.contains("\n") || 
                          stringValue.contains("\r")
        
        if needsQuoting {
            // Escape quotes by doubling them
            let escapedValue = stringValue.replacingOccurrences(of: "\"", with: "\"\"")
            return "\"\(escapedValue)\""
        } else {
            return stringValue
        }
    }
    
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
        
        // Check for date (simple format: yyyy-MM-dd)
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
    
    private static func parseCSVLine(_ line: String, separator: String) -> [String] {
        var result: [String] = []
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
            } else if char == Character(separator) && !inQuotes {
                result.append(currentField)
                currentField = ""
            } else {
                currentField += String(char)
            }
            
            i += 1
        }
        
        result.append(currentField)
        return result
    }
}

// MARK: - String Extension for Pattern Matching

private extension String {
    func matches(pattern: String) -> Bool {
        return range(of: pattern, options: .regularExpression) != nil
    }
}

// (Insert CSVUtils and related helpers here, as previously defined) 