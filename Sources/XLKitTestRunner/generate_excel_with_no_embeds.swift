//
//  generate_excel_with_no_embeds.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit

// MARK: - Configuration

// CSV file name
let csvFileName = "Embed-Test.csv"
let outputExcelFile = "Test-Workflows/Embed-Test.xlsx"

// Dynamically find the test data folder and CSV file
func findCSVFile() -> String? {
    let currentDir = FileManager.default.currentDirectoryPath
    let possiblePaths = [
        "Test-Data",
        "test-data", 
        "TestData",
        "testdata"
    ]
    
    for path in possiblePaths {
        let fullPath = "\(currentDir)/\(path)"
        if FileManager.default.fileExists(atPath: fullPath) {
            // Search for CSV file in this directory and subdirectories
            let fileManager = FileManager.default
            do {
                let enumerator = fileManager.enumerator(atPath: fullPath)
                while let filePath = enumerator?.nextObject() as? String {
                    if filePath.hasSuffix(".csv") && filePath.contains(csvFileName) {
                        let csvPath = "\(fullPath)/\(filePath)"
                        print("[INFO] Found CSV file: \(csvPath)")
                        return csvPath
                    }
                }
            } catch {
                print("[WARNING] Error searching in \(fullPath): \(error)")
            }
        }
    }
    
    // Try parent directory
    let parentDir = (currentDir as NSString).deletingLastPathComponent
    for path in possiblePaths {
        let fullPath = "\(parentDir)/\(path)"
        if FileManager.default.fileExists(atPath: fullPath) {
            // Search for CSV file in this directory and subdirectories
            let fileManager = FileManager.default
            do {
                let enumerator = fileManager.enumerator(atPath: fullPath)
                while let filePath = enumerator?.nextObject() as? String {
                    if filePath.hasSuffix(".csv") && filePath.contains(csvFileName) {
                        let csvPath = "\(fullPath)/\(filePath)"
                        print("[INFO] Found CSV file: \(csvPath)")
                        return csvPath
                    }
                }
            } catch {
                print("[WARNING] Error searching in \(fullPath): \(error)")
            }
        }
    }
    
    return nil
}

// Find CSV file
guard let csvFilePath = findCSVFile() else {
    print("[ERROR] CSV file '\(csvFileName)' not found. Searched in:")
    print("  - Test-Data (and subdirectories)")
    print("  - test-data (and subdirectories)")
    print("  - TestData (and subdirectories)")
    print("  - testdata (and subdirectories)")
    print("  - And parent directories")
    exit(1)
}

print("[INFO] Using CSV file: \(csvFilePath)")

// MARK: - CSV Parsing
func parseCSV(_ csv: String) -> (rows: [[String]], headers: [String]) {
    let lines = csv.components(separatedBy: .newlines).filter { !$0.isEmpty }
    guard let headerLine = lines.first else { return ([], []) }
    let headers = headerLine.components(separatedBy: ",")
    let rows = lines.dropFirst().map { $0.components(separatedBy: ",") }
    return (rows, headers)
}

guard let csvData = try? String(contentsOfFile: csvFilePath, encoding: .utf8) else {
    print("[ERROR] Failed to read CSV file: \(csvFilePath)")
    exit(1)
}

let (rows, headers) = parseCSV(csvData)
print("[INFO] Parsed CSV with \(headers.count) columns and \(rows.count) data rows")

// MARK: - Create Excel File with XLKit
print("[INFO] Creating Excel workbook...")
let workbook = XLKit.createWorkbook()
let sheet = workbook.addSheet(name: "Embed Test")

// Write headers
print("[INFO] Writing headers...")
for (col, header) in headers.enumerated() {
    sheet.setCell(row: 1, column: col + 1, value: .string(header))
}

// Write data rows
print("[INFO] Writing data rows...")
for (rowIdx, row) in rows.enumerated() {
    let excelRow = rowIdx + 2
    for (colIdx, value) in row.enumerated() {
        sheet.setCell(row: excelRow, column: colIdx + 1, value: .string(value))
    }
}

// Save the workbook
print("[INFO] Saving Excel file...")
do {
    try XLKit.saveWorkbook(workbook, to: URL(fileURLWithPath: outputExcelFile))
    print("[SUCCESS] Excel file created: \(outputExcelFile)")
} catch {
    print("[ERROR] Failed to save Excel file: \(error)")
    exit(1)
} 