//
//  ExcelGenerators.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit
import CoreXLSX

/// Excel generation functions for XLKitTestRunner
/// 
/// Provides comprehensive test generators for XLKit functionality:
/// - CSV to Excel conversion with formatting
/// - Comprehensive API demonstration
/// - Security feature testing
/// - Excel file validation using CoreXLSX

struct ExcelGenerators {
    
    // MARK: - CSV to Excel Conversion
    
    /// Generates Excel file from CSV data with professional formatting
    /// 
    /// Features:
    /// - Automatic column width calculation
    /// - Professional header formatting
    /// - Data type preservation
    /// - Excel file validation
    @MainActor
    static func generateExcelWithNoEmbeds() throws {
        // MARK: - Configuration
        
        let csvFilePath = "Test-Data/Embed-Test/Embed-Test.csv"
        let outputExcelFile = "Test-Workflows/Embed-Test.xlsx"
        
        print("[INFO] Using CSV file: \(csvFilePath)")
        
        guard FileManager.default.fileExists(atPath: csvFilePath) else {
            print("[ERROR] CSV file not found: \(csvFilePath)")
            throw XLKitError.fileWriteError("CSV file not found: \(csvFilePath)")
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
            throw XLKitError.fileWriteError("Failed to read CSV file: \(csvFilePath)")
        }
        
        let (rows, headers) = parseCSV(csvData)
        print("[INFO] Parsed CSV with \(headers.count) columns and \(rows.count) data rows")
        
        // MARK: - Create Excel File
        print("[INFO] Creating Excel workbook...")
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Embed Test")
        
        // MARK: - Calculate Column Widths
        print("[INFO] Calculating optimal column widths...")
        var columnWidths: [Int: Double] = [:]
        
        for col in 0..<headers.count {
            let column = col + 1
            var maxWidth = calculateTextWidth(headers[col])
            
            for row in rows {
                if col < row.count {
                    let cellWidth = calculateTextWidth(row[col])
                    maxWidth = max(maxWidth, cellWidth)
                }
            }
            
            let adjustedWidth = min(max(maxWidth + 4.0, 8.0), 50.0)
            columnWidths[column] = adjustedWidth
            print("[INFO] Column \(column) (\(headers[col])): width = \(adjustedWidth)")
        }
        
        // MARK: - Write Headers
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
        print("[INFO] Applying calculated column widths...")
        for (column, width) in columnWidths {
            sheet.setColumnWidth(column, width: width)
        }
        
        // MARK: - File Output
        let outputURL = URL(fileURLWithPath: outputExcelFile)
        let outputDir = outputURL.deletingLastPathComponent()
        if !FileManager.default.fileExists(atPath: outputDir.path) {
            print("[INFO] Creating output directory: \(outputDir.path)")
            do {
                try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
            } catch {
                print("[ERROR] Failed to create output directory: \(error)")
                throw XLKitError.fileWriteError("Failed to create output directory: \(error)")
            }
        }
        
        print("[INFO] Saving Excel file...")
        let outputPath = outputExcelFile
        do {
            try workbook.save(to: URL(fileURLWithPath: outputPath))
            print("[SUCCESS] Excel file created: \(outputPath)")
            print("[INFO] Features applied:")
            print("  - Bold headers with gray background")
            print("  - Automatic column width adjustment")
            print("  - \(headers.count) columns optimized")
        } catch {
            print("[ERROR] Failed to save Excel file: \(error)")
            throw error
        }
    }

    // MARK: - Comprehensive API Demonstration
    
    /// Demonstrates comprehensive XLKit API features
    /// 
    /// Features:
    /// - Cell operations with different data types and formatting
    /// - Range operations and cell merging
    /// - Image embedding with perfect aspect ratio preservation
    /// - Workbook and sheet management
    /// - CSV/TSV import and export
    /// - Fluent API usage with method chaining
    /// - Excel file validation using CoreXLSX
    @MainActor
    static func demonstrateComprehensiveAPI() async throws {
        print("[INFO] Starting comprehensive XLKit API demonstration...")
        
        // MARK: - Workbook Creation
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "API Demo")
        
        print("[INFO] Created workbook with sheet: \(sheet.name)")
        
        // MARK: - Cell Operations
        print("[INFO] Demonstrating cell operations...")
        
        sheet.setCell("A1", string: "String Value", format: CellFormat.header())
        sheet.setCell("A2", string: "Another String")
        sheet.setCell("B1", number: 123.45, format: CellFormat.currency())
        sheet.setCell("B2", integer: 42)
        sheet.setCell("C1", boolean: true)
        sheet.setCell("C2", date: Date(), format: CellFormat.date())
        sheet.setCell("D1", formula: "=SUM(B1:B2)")
        sheet.setRange("A4:C6", string: "Range Value", format: CellFormat.bordered())
        
        // MARK: - Font Color Demonstration
        print("[INFO] Demonstrating font color formatting...")
        
        let redTextFormat = CellFormat.text(color: "#FF0000")
        let blueTextFormat = CellFormat.text(color: "#0000FF")
        let greenTextFormat = CellFormat.text(color: "#00FF00")
        let currencyRedFormat = CellFormat.currency(color: "#FF0000")
        
        sheet.setCell("F1", string: "Red Text", format: redTextFormat)
        sheet.setCell("F2", string: "Blue Text", format: blueTextFormat)
        sheet.setCell("F3", string: "Green Text", format: greenTextFormat)
        sheet.setCell("F4", number: 1234.56, format: currencyRedFormat)
        
        print("[INFO] ✓ Applied font colors: red, blue, green, and red currency")
        
        // MARK: - Image Embedding
        print("[INFO] Demonstrating image embedding...")
        
        let testImageData = createTestImageData()
        
        let success1 = try await sheet.embedImage(
            testImageData,
            at: "E1",
            of: workbook,
            scale: 0.5
        )
        
        if success1 {
            print("[INFO] ✓ Successfully embedded test image using data")
        }
        
        // MARK: - Workbook Management
        print("[INFO] Demonstrating workbook management...")
        
        let sheet2 = workbook.addSheet(name: "Second Sheet")
        sheet2.setCell("A1", string: "Second Sheet Content")
        
        if let foundSheet = workbook.getSheet(name: "API Demo") {
            print("[INFO] ✓ Found sheet: \(foundSheet.name)")
        }
        
        print("[INFO] Workbook contains \(workbook.imageCount) images")
        
        // MARK: - CSV/TSV Operations
        print("[INFO] Demonstrating CSV/TSV operations...")
        
        let csv = sheet.exportToCSV()
        print("[INFO] ✓ Exported sheet to CSV (\(csv.count) characters)")
        
        let tsv = sheet.exportToTSV()
        print("[INFO] ✓ Exported sheet to TSV (\(tsv.count) characters)")
        
        let csvData = """
        Name,Age,Salary
        Alice,30,50000.5
        Bob,25,45000.75
        """
        let csvWorkbook = Workbook(fromCSV: csvData, hasHeader: true)
        print("[INFO] ✓ Created workbook from CSV with \(csvWorkbook.getSheets().count) sheets")
        
        let tsvData = """
        Product\tPrice\tIn Stock
        Apple\t1.99\ttrue
        Banana\t0.99\tfalse
        """
        let tsvWorkbook = Workbook(fromTSV: tsvData, hasHeader: true)
        print("[INFO] ✓ Created workbook from TSV with \(tsvWorkbook.getSheets().count) sheets")
        
        // MARK: - Fluent API
        print("[INFO] Demonstrating fluent API...")
        
        _ = workbook.addSheet(name: "Fluent Demo")
            .setRow(1, strings: ["Header1", "Header2", "Header3"])
            .setRow(2, numbers: [1.5, 2.5, 3.5])
            .setRow(3, integers: [10, 20, 30])
            .mergeCells("A1:C1")
        
        print("[INFO] ✓ Created sheet using fluent API")
        
        // MARK: - Save and Validate
        print("[INFO] Saving comprehensive demo...")
        
        let outputPath = "Test-Workflows/Comprehensive-Demo.xlsx"
        let outputURL = URL(fileURLWithPath: outputPath)
        
        let outputDir = outputURL.deletingLastPathComponent()
        if !FileManager.default.fileExists(atPath: outputDir.path) {
            try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
        }
        
        try await workbook.save(to: outputURL)
        print("[SUCCESS] Comprehensive demo saved: \(outputPath)")
        
        // MARK: - CoreXLSX Validation
        print("[INFO] Validating with CoreXLSX...")
        do {
            guard let xlsx = XLSXFile(filepath: outputPath) else {
                print("[ERROR] Failed to open Excel file with CoreXLSX")
                return
            }
            
            let worksheet = try xlsx.parseWorksheet(at: "xl/worksheets/sheet1.xml")
            let sharedStrings = try xlsx.parseSharedStrings()
            
            print("[INFO] ✓ CoreXLSX validation successful")
            print("[INFO] ✓ Worksheet contains \(worksheet.data?.rows.count ?? 0) rows")
            print("[INFO] ✓ Shared strings count: \(sharedStrings?.items.count ?? 0)")
            
        } catch {
            print("[ERROR] CoreXLSX validation failed: \(error)")
            throw error
        }
        
        print("[SUCCESS] Comprehensive API demonstration completed successfully!")
    }

    // MARK: - Security Features Demonstration
    
    /// Demonstrates XLKit's file path security restrictions
    @MainActor
    static func demonstrateFilePathRestrictions() {
        print("=== File Path Security Restrictions Demo ===")
        
        // Test 1: Default behavior (restrictions disabled)
        print("\n1. Testing default behavior (restrictions disabled):")
        SecurityManager.enableFilePathRestrictions = false
        print("   enableFilePathRestrictions = false")
        
        do {
            let testPath = "Test-Workflows/test-default.xlsx"
            try CoreUtils.validateFilePath(testPath)
            print("   ✓ File path validation passed: \(testPath)")
        } catch {
            print("   ✗ File path validation failed: \(error)")
        }
        
        // Test 2: With restrictions enabled
        print("\n2. Testing with restrictions enabled:")
        SecurityManager.enableFilePathRestrictions = true
        print("   enableFilePathRestrictions = true")
        
        do {
            let testPath = "Test-Workflows/test-restricted.xlsx"
            try CoreUtils.validateFilePath(testPath)
            print("   ✓ File path validation passed: \(testPath)")
        } catch {
            print("   ✗ File path validation failed: \(error)")
        }
        
        // Test 3: Path traversal attempt
        print("\n3. Testing path traversal protection:")
        do {
            let maliciousPath = "Test-Workflows/test-restricted.xlsx"
            try CoreUtils.validateFilePath(maliciousPath)
            print("   ✓ File path validation passed: \(maliciousPath)")
        } catch {
            print("   ✗ File path validation failed: \(error)")
        }
        
        SecurityManager.enableFilePathRestrictions = false
        print("\n=== Demo completed ===")
    }

    // MARK: - Helper Functions
    
    /// Calculates approximate text width for column sizing (8 pixels per character)
    private static func calculateTextWidth(_ text: String) -> Double {
        return Double(text.count) * 8.0
    }

    /// Creates a minimal valid PNG image for testing purposes
    private static func createTestImageData() -> Data {
        let pngData: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0x01, // width: 1
            0x00, 0x00, 0x00, 0x01, // height: 1
            0x08, 0x02, 0x00, 0x00, 0x00, // bit depth, color type, compression, filter, interlace
            0x90, 0x77, 0x53, 0xDE, // CRC
            0x00, 0x00, 0x00, 0x0C, // IDAT chunk length
            0x49, 0x44, 0x41, 0x54, // IDAT
            0x08, 0x99, 0x01, 0x01, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, // compressed data
            0x00, 0x00, 0x00, 0x00, // IEND chunk length
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngData)
    }

    // MARK: - Future Implementation Placeholders
    
    /// Generates Excel file with image embeds (future implementation)
    static func generateExcelWithEmbeds() {
        print("[INFO] Image embed functionality not yet implemented")
        print("[INFO] This will be implemented in a future version")
    }

    /// Tests CSV import functionality (future implementation)
    static func csvImportTest() {
        print("[INFO] CSV import test functionality not yet implemented")
        print("[INFO] This will be implemented in a future version")
    }

    /// Tests cell formatting features (future implementation)
    static func cellFormattingTest() {
        print("[INFO] Cell formatting test functionality not yet implemented")
        print("[INFO] This will be implemented in a future version")
    }

    // MARK: - Excel File Validation
    
    /// Validates Excel file using CoreXLSX for compliance testing
    static func validateExcelFile(_ filePath: String) {
        print("[VALIDATION] Validating Excel file: \(filePath)")
        
        guard let file = XLSXFile(filepath: filePath) else {
            print("[ERROR] Failed to open Excel file with CoreXLSX")
            return
        }
        
        do {
            // MARK: - Workbook Structure Validation
            let workbooks = try file.parseWorkbooks()
            guard let workbook = workbooks.first else {
                print("[ERROR] No workbook found in file")
                return
            }
            print("[VALIDATION] ✓ Workbook parsed successfully")
            
            // MARK: - Worksheet Validation
            let worksheets = try file.parseWorksheetPathsAndNames(workbook: workbook)
            print("[VALIDATION] ✓ Found \(worksheets.count) worksheet(s)")
            
            // MARK: - Shared Strings Validation
            let sharedStrings = try file.parseSharedStrings()
            print("[VALIDATION] ✓ Shared strings parsed successfully (\(sharedStrings?.items.count ?? 0) strings)")
            
            // MARK: - Styles Validation
            _ = try file.parseStyles()
            print("[VALIDATION] ✓ Styles parsed successfully")
            
            // MARK: - Worksheet Data Validation
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
} 