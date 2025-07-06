//
//  main.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit

// MARK: - Configuration

// Explicit path to the test data folder and CSV file
let testDataFolder = "Test-Data/Embed-Test 2025-07-06 2-25-34 PM"
let csvFileName = "Embed-Test.csv"
let outputExcelFile = "Test-Workflows/Embed-Test-with-Images.xlsx"

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

// DEBUG: Print CSV parsing results
print("[DEBUG] CSV Headers: \(headers)")
print("[DEBUG] Number of rows parsed: \(rows.count)")
if !rows.isEmpty {
    print("[DEBUG] First row: \(rows[0])")
}

// Find the index of the 'Image Filename' column
let imageFilenameCol = headers.firstIndex(of: "Image Filename")
if imageFilenameCol == nil {
    print("[ERROR] 'Image Filename' column not found in CSV header")
    exit(1)
}
let imageFilenameIndex = imageFilenameCol!

// Prepare new headers with 'Image' column after 'Image Filename'
var newHeaders = headers
newHeaders.insert("Image", at: imageFilenameIndex + 1)

// Prepare new rows with an empty cell for the image column
var newRows: [[String]] = []
for row in rows {
    var newRow = row
    // Insert empty string for the image column
    if newRow.count > imageFilenameIndex {
        newRow.insert("", at: imageFilenameIndex + 1)
    } else {
        // Pad row if it's short
        while newRow.count <= imageFilenameIndex { newRow.append("") }
        newRow.append("")
    }
    newRows.append(newRow)
}

// DEBUG: Print processed data
print("[DEBUG] New headers: \(newHeaders)")
print("[DEBUG] Number of processed rows: \(newRows.count)")
if !newRows.isEmpty {
    print("[DEBUG] First processed row: \(newRows[0])")
}

// MARK: - Create Excel File with XLKit
print("[INFO] Creating Excel workbook...")
let workbook = XLKit.createWorkbook()
let sheet = workbook.addSheet(name: "Embed Test")

// Write headers
print("[INFO] Writing headers...")
for (col, header) in newHeaders.enumerated() {
    sheet.setCell(row: 1, column: col + 1, value: .string(header))
    print("[DEBUG] Wrote header '\(header)' at column \(col + 1)")
}

// Write data rows (temporarily without images)
print("[INFO] Writing data rows...")
for (rowIdx, row) in newRows.enumerated() {
    let excelRow = rowIdx + 2
    print("[DEBUG] Processing row \(excelRow): \(row)")
    for (colIdx, value) in row.enumerated() {
        // Temporarily skip image embedding to avoid drawing XML issues
        if colIdx == imageFilenameIndex + 1 {
            // Just write a placeholder for now
            sheet.setCell(row: excelRow, column: colIdx + 1, value: .string("[Image Placeholder]"))
            print("[DEBUG] Wrote image placeholder at row \(excelRow), column \(colIdx + 1)")
        } else {
            sheet.setCell(row: excelRow, column: colIdx + 1, value: .string(value))
            print("[DEBUG] Wrote value '\(value)' at row \(excelRow), column \(colIdx + 1)")
        }
    }
}

// DEBUG: Print sheet contents before saving
print("[DEBUG] Sheet used cells: \(sheet.getUsedCells())")

// Save the workbook
print("[INFO] Saving Excel file...")
do {
    try XLKit.saveWorkbook(workbook, to: URL(fileURLWithPath: outputExcelFile))
    print("[SUCCESS] Excel file created: \(outputExcelFile)")
} catch {
    print("[ERROR] Failed to save Excel file: \(error)")
    exit(1)
} 