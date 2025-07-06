//
//  ExcelGenerators.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit

// MARK: - Excel Generation Functions

/// Generates an Excel file from CSV data without image embeds
func generateExcelWithNoEmbeds() {
    // MARK: - Configuration
    
    // Fixed CSV file path
    let csvFilePath = "Test-Data/Embed-Test/Embed-Test.csv"
    let outputExcelFile = "Test-Workflows/Embed-Test.xlsx"
    
    print("[INFO] Using CSV file: \(csvFilePath)")
    
    // Check if CSV file exists
    guard FileManager.default.fileExists(atPath: csvFilePath) else {
        print("[ERROR] CSV file not found: \(csvFilePath)")
        exit(1)
    }
    
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
    
    // Ensure output directory exists
    let outputURL = URL(fileURLWithPath: outputExcelFile)
    let outputDir = outputURL.deletingLastPathComponent()
    if !FileManager.default.fileExists(atPath: outputDir.path) {
        print("[INFO] Creating output directory: \(outputDir.path)")
        do {
            try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
        } catch {
            print("[ERROR] Failed to create output directory: \(error)")
            exit(1)
        }
    }
    
    // Save the workbook
    print("[INFO] Saving Excel file...")
    do {
        try XLKit.saveWorkbook(workbook, to: outputURL)
        print("[SUCCESS] Excel file created: \(outputExcelFile)")
    } catch {
        print("[ERROR] Failed to save Excel file: \(error)")
        exit(1)
    }
}

/// Generates an Excel file with image embeds (future implementation)
func generateExcelWithEmbeds() {
    print("[INFO] Image embed functionality not yet implemented")
    print("[INFO] This will be implemented in a future version")
    exit(0)
}

/// Tests CSV import functionality (future implementation)
func csvImportTest() {
    print("[INFO] CSV import test functionality not yet implemented")
    print("[INFO] This will be implemented in a future version")
    exit(0)
}

/// Tests cell formatting features (future implementation)
func cellFormattingTest() {
    print("[INFO] Cell formatting test functionality not yet implemented")
    print("[INFO] This will be implemented in a future version")
    exit(0)
} 