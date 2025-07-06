#!/usr/bin/env swift

import Foundation
// import XLKit // Uncomment this line if running in a real project with XLKit available

// MARK: - Configuration
let testDataPath = "Test-Data/Embed-Test 2025-07-06 2-25-34 PM"
let csvFile = "\(testDataPath)/Embed-Test.csv"
let outputExcelFile = "Embed-Test.xlsx"

// MARK: - CSV Parsing
func parseCSV(_ csv: String) -> ([[String]], [String]) {
    let lines = csv.components(separatedBy: .newlines).filter { !$0.isEmpty }
    guard let header = lines.first else { return ([], []) }
    let headers = header.components(separatedBy: ",")
    let rows = lines.dropFirst().map { $0.components(separatedBy: ",") }
    return (rows, headers)
}

// MARK: - Main Logic
print("[INFO] Reading CSV data from \(csvFile)")
guard let csvData = try? String(contentsOfFile: csvFile) else {
    print("[ERROR] Failed to read CSV file at \(csvFile)")
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

// MARK: - Create Excel File (Pseudocode, replace with XLKit API)
// let workbook = XLKit.createWorkbook()
// let sheet = workbook.addSheet(name: "Embed Test")
//
// // Write headers
// for (col, header) in newHeaders.enumerated() {
//     sheet.setCell(row: 1, column: col + 1, value: .string(header))
// }
//
// // Write data rows and embed images
// for (rowIdx, row) in newRows.enumerated() {
//     let excelRow = rowIdx + 2
//     for (colIdx, value) in row.enumerated() {
//         // If this is the 'Image' column, embed the image
//         if colIdx == imageFilenameIndex + 1, row.count > imageFilenameIndex {
//             let imageFilename = row[imageFilenameIndex].trimmingCharacters(in: .whitespacesAndNewlines)
//             let imagePath = "\(testDataPath)/\(imageFilename)"
//             if FileManager.default.fileExists(atPath: imagePath) {
//                 let imageURL = URL(fileURLWithPath: imagePath)
//                 try? sheet.addImage(imageURL, at: CellCoordinate(row: excelRow, column: colIdx + 1).excelAddress)
//                 sheet.autoSizeColumn(colIdx + 1, forImageAt: CellCoordinate(row: excelRow, column: colIdx + 1).excelAddress)
//             }
//         } else {
//             sheet.setCell(row: excelRow, column: colIdx + 1, value: .string(value))
//         }
//     }
// }
//
// try XLKit.saveWorkbook(workbook, to: URL(fileURLWithPath: outputExcelFile))
// print("[INFO] Excel file created: \(outputExcelFile)")

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