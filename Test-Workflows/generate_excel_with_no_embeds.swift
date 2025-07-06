//
//  generate_excel_with_no_embeds.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit

// MARK: - Configuration

// Explicit path to the test data folder and CSV file
let testDataFolder = "Test-Data/Embed-Test 2025-07-06 2-25-34 PM"
let csvFileName = "Embed-Test.csv"
let outputExcelFile = "Test-Workflows/Embed-Test.xlsx"

// Resolve absolute paths for local and CI environments
let currentDir = FileManager.default.currentDirectoryPath
let testDataPath: String
if FileManager.default.fileExists(atPath: "\(currentDir)/\(testDataFolder)") {
    testDataPath = "\(currentDir)/\(testDataFolder)"
} else if FileManager.default.fileExists(atPath: testDataFolder) {
    testDataPath = testDataFolder
} else {
    print("[ERROR] Test data folder not found: \(testDataFolder)")
    exit(1)
}

let csvFilePath = "\(testDataPath)/\(csvFileName)"
if !FileManager.default.fileExists(atPath: csvFilePath) {
    print("[ERROR] CSV file not found: \(csvFilePath)")
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