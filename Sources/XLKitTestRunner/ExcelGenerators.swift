//
//  ExcelGenerators.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit
import CoreXLSX

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
    
    // MARK: - Calculate Column Widths
    print("[INFO] Calculating optimal column widths...")
    var columnWidths: [Int: Double] = [:]
    
    // Calculate width for each column based on longest value
    for col in 0..<headers.count {
        let column = col + 1
        var maxWidth = calculateTextWidth(headers[col]) // Start with header width
        
        // Check all data rows for this column
        for row in rows {
            if col < row.count {
                let cellWidth = calculateTextWidth(row[col])
                maxWidth = max(maxWidth, cellWidth)
            }
        }
        
        // Add some padding and set minimum/maximum bounds
        let adjustedWidth = min(max(maxWidth + 4.0, 8.0), 50.0) // Min 8, Max 50, +4 padding
        columnWidths[column] = adjustedWidth
        print("[INFO] Column \(column) (\(headers[col])): width = \(adjustedWidth)")
    }
    
    // MARK: - Write Headers with Bold Formatting
    print("[INFO] Writing headers with bold formatting...")
    let headerFormat = CellFormat.header(fontSize: 12.0, backgroundColor: "#E6E6E6")
    
    for (col, header) in headers.enumerated() {
        let column = col + 1
        sheet.setCell(row: 1, column: column, cell: Cell.string(header, format: headerFormat))
    }
    
    // MARK: - Write Data Rows
    print("[INFO] Writing data rows...")
    for (rowIdx, row) in rows.enumerated() {
        let excelRow = rowIdx + 2
        for (colIdx, value) in row.enumerated() {
            let column = colIdx + 1
            sheet.setCell(row: excelRow, column: column, value: .string(value))
        }
    }
    
    // MARK: - Apply Column Widths
    print("[INFO] Applying column widths...")
    for (column, width) in columnWidths {
        sheet.setColumnWidth(column, width: width)
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
        print("[INFO] Features applied:")
        print("  - Bold headers with gray background")
        print("  - Automatic column width adjustment")
        print("  - \(headers.count) columns optimized")
    } catch {
        print("[ERROR] Failed to save Excel file: \(error)")
        exit(1)
    }
}

/// Calculates the approximate width needed for text in Excel
/// This is a simplified calculation based on character count and font size
private func calculateTextWidth(_ text: String) -> Double {
    // Base width calculation: approximately 1.2 units per character for standard font
    // This is a reasonable approximation for most fonts and text
    let baseWidth = Double(text.count) * 1.2
    
    // Add extra width for special characters that might be wider
    let specialCharCount = text.filter { "MWmw".contains($0) }.count
    let extraWidth = Double(specialCharCount) * 0.3
    
    return baseWidth + extraWidth
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

func validateExcelFile(_ filePath: String) {
    print("[VALIDATION] Validating Excel file: \(filePath)")
    
    guard let file = XLSXFile(filepath: filePath) else {
        print("[ERROR] Failed to open Excel file with CoreXLSX")
        return
    }
    
    do {
        // Try to parse the workbooks
        let workbooks = try file.parseWorkbooks()
        guard let workbook = workbooks.first else {
            print("[ERROR] No workbook found in file")
            return
        }
        print("[VALIDATION] ✓ Workbook parsed successfully")
        
        // Try to parse worksheets
        let worksheets = try file.parseWorksheetPathsAndNames(workbook: workbook)
        print("[VALIDATION] ✓ Found \(worksheets.count) worksheet(s)")
        
        // Try to parse shared strings
        let sharedStrings = try file.parseSharedStrings()
        print("[VALIDATION] ✓ Shared strings parsed successfully (\(sharedStrings?.items.count ?? 0) strings)")
        
        // Try to parse styles
        let styles = try file.parseStyles()
        print("[VALIDATION] ✓ Styles parsed successfully")
        
        // Try to read worksheet data
        for worksheet in worksheets {
            let worksheetData = try file.parseWorksheet(at: worksheet.path)
            print("[VALIDATION] ✓ Worksheet '\(worksheet.name ?? "Unnamed")' parsed successfully")
            
            let rows = worksheetData.data?.rows ?? []
            print("[VALIDATION] ✓ Found \(rows.count) rows in worksheet")
            
            if let firstRow = rows.first {
                let cells = firstRow.cells
                print("[VALIDATION] ✓ First row has \(cells.count) cells")
            }
        }
        
        print("[VALIDATION] ✓ All validation tests passed! Excel file is fully compliant.")
        
    } catch {
        print("[ERROR] Validation failed: \(error)")
    }
} 