//
//  generate_excel_with_embeds.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

#!/usr/bin/env swift

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

// MARK: - Create Excel File with XLKit
print("[INFO] Creating Excel workbook...")
let workbook = XLKit.createWorkbook()
let sheet = workbook.addSheet(name: "Embed Test")

// Write headers
print("[INFO] Writing headers...")
for (col, header) in newHeaders.enumerated() {
    sheet.setCell(row: 1, column: col + 1, value: .string(header))
}

// Write data rows and embed images
print("[INFO] Writing data rows and embedding images...")
for (rowIdx, row) in newRows.enumerated() {
    let excelRow = rowIdx + 2
    for (colIdx, value) in row.enumerated() {
        // If this is the 'Image' column, embed the image
        if colIdx == imageFilenameIndex + 1, row.count > imageFilenameIndex {
            let imageFilename = row[imageFilenameIndex].trimmingCharacters(in: .whitespacesAndNewlines)
            let imagePath = "\(testDataPath)/\(imageFilename)"
            if FileManager.default.fileExists(atPath: imagePath) {
                let imageURL = URL(fileURLWithPath: imagePath)
                print("[INFO] Embedding image: \(imageFilename) at row \(excelRow)")
                do {
                    try sheet.addImage(from: imageURL, at: "B\(excelRow)")
                    sheet.autoSizeColumn(2, forImageAt: "B\(excelRow)")
                } catch {
                    print("[WARNING] Failed to embed image \(imageFilename): \(error)")
                }
            } else {
                print("[WARNING] Image file not found: \(imagePath)")
            }
        } else {
            sheet.setCell(row: excelRow, column: colIdx + 1, value: .string(value))
        }
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

// MARK: - Fallback: CSV to XLSX (if XLKit not available)
// For demonstration, just write the new CSV (no images)
let tempCSV = ([newHeaders] + newRows).map { $0.joined(separator: ",") }.joined(separator: "\n")
let tempCSVFile = "Embed-Test-with-Image-Column.csv"
do {
    try tempCSV.write(toFile: tempCSVFile, atomically: true, encoding: .utf8)
    print("[INFO] Wrote fallback CSV with 'Image' column: \(tempCSVFile)")
    print("[INFO] To enable Excel+image output, uncomment and use the XLKit code in this script.")
} catch {
    print("[ERROR] Failed to write fallback CSV: \(error)")
    exit(1)
} 