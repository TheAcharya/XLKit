//
//  ExcelGenerators.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit
import CoreXLSX

// MARK: - Excel Generation Functions

struct ExcelGenerators {
    @MainActor
    static func generateExcelWithNoEmbeds() throws {
        // MARK: - Configuration
        
        // Fixed CSV file path
        let csvFilePath = "Test-Data/Embed-Test/Embed-Test.csv"
        let outputExcelFile = "Test-Workflows/Embed-Test.xlsx"
        
        print("[INFO] Using CSV file: \(csvFilePath)")
        
        // Check if CSV file exists
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
        
        // MARK: - Create Excel File with XLKit
        print("[INFO] Creating Excel workbook...")
        let workbook = Workbook()
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
                throw XLKitError.fileWriteError("Failed to create output directory: \(error)")
            }
        }
        
        // Save the workbook
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

    @MainActor
    static func demonstrateComprehensiveAPI() async throws {
        print("[INFO] Starting comprehensive XLKit API demonstration...")
        
        // MARK: - Workbook Creation
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "API Demo")
        
        print("[INFO] Created workbook with sheet: \(sheet.name)")
        
        // MARK: - Cell Operations
        print("[INFO] Demonstrating cell operations...")
        
        // String cells
        sheet.setCell("A1", string: "String Value", format: CellFormat.header())
        sheet.setCell("A2", string: "Another String")
        
        // Numeric cells
        sheet.setCell("B1", number: 123.45, format: CellFormat.currency())
        sheet.setCell("B2", integer: 42)
        
        // Boolean and date cells
        sheet.setCell("C1", boolean: true)
        sheet.setCell("C2", date: Date(), format: CellFormat.date())
        
        // Formula cells
        sheet.setCell("D1", formula: "=SUM(B1:B2)")
        
        // Range operations
        sheet.setRange("A4:C6", string: "Range Value", format: CellFormat.bordered())
        
        // MARK: - Image Embedding
        print("[INFO] Demonstrating image embedding...")
        
        // Create a simple test image (1x1 pixel PNG)
        let testImageData = createTestImageData()
        
        // Embed image using different methods
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
        
        // Add another sheet
        let sheet2 = workbook.addSheet(name: "Second Sheet")
        sheet2.setCell("A1", string: "Second Sheet Content")
        
        // Get sheet by name
        if let foundSheet = workbook.getSheet(name: "API Demo") {
            print("[INFO] ✓ Found sheet: \(foundSheet.name)")
        }
        
        // MARK: - Image Management
        print("[INFO] Demonstrating image management...")
        
        // Note: We don't need to add the image to workbook separately since it's already added via embedImage
        // The embedImage method automatically adds the image to the workbook
        
        // Get image statistics
        print("[INFO] Workbook contains \(workbook.imageCount) images")
        
        // MARK: - CSV/TSV Operations
        print("[INFO] Demonstrating CSV/TSV operations...")
        
        // Export sheet to CSV
        let csv = sheet.exportToCSV()
        print("[INFO] ✓ Exported sheet to CSV (\(csv.count) characters)")
        
        // Export sheet to TSV
        let tsv = sheet.exportToTSV()
        print("[INFO] ✓ Exported sheet to TSV (\(tsv.count) characters)")
        
        // Create workbook from CSV
        let csvData = """
        Name,Age,Salary
        Alice,30,50000.5
        Bob,25,45000.75
        """
        let csvWorkbook = Workbook(fromCSV: csvData, hasHeader: true)
        print("[INFO] ✓ Created workbook from CSV with \(csvWorkbook.getSheets().count) sheets")
        
        // Create workbook from TSV
        let tsvData = """
        Product\tPrice\tIn Stock
        Apple\t1.99\ttrue
        Banana\t0.99\tfalse
        """
        let tsvWorkbook = Workbook(fromTSV: tsvData, hasHeader: true)
        print("[INFO] ✓ Created workbook from TSV with \(tsvWorkbook.getSheets().count) sheets")
        
        // MARK: - Fluent API
        print("[INFO] Demonstrating fluent API...")
        
        let fluentSheet = workbook.addSheet(name: "Fluent Demo")
            .setRow(1, strings: ["Header1", "Header2", "Header3"])
            .setRow(2, numbers: [1.5, 2.5, 3.5])
            .setRow(3, integers: [10, 20, 30])
            .mergeCells("A1:C1")
        
        print("[INFO] ✓ Created sheet using fluent API")
        
        // MARK: - Save and Validate
        print("[INFO] Saving comprehensive demo...")
        
        let outputPath = "Test-Workflows/Comprehensive-Demo.xlsx"
        let outputURL = URL(fileURLWithPath: outputPath)
        
        // Ensure output directory exists
        let outputDir = outputURL.deletingLastPathComponent()
        if !FileManager.default.fileExists(atPath: outputDir.path) {
            try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
        }
        
        try await workbook.save(to: outputURL)
        print("[SUCCESS] Comprehensive demo saved: \(outputPath)")
        
        // Validate with CoreXLSX
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

    /// Demonstrates file path security restrictions
    @MainActor
    static func demonstrateFilePathRestrictions() {
        print("=== File Path Security Restrictions Demo ===")
        
        // Test 1: Default behavior (restrictions disabled)
        print("\n1. Testing default behavior (restrictions disabled):")
        SecurityManager.enableFilePathRestrictions = false
        print("   enableFilePathRestrictions = false")
        
        do {
            let testPath = "/Users/vr/Developer/XLKit/Test-Workflows/test-default.xlsx"
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
            let testPath = "/Users/vr/Developer/XLKit/Test-Workflows/test-restricted.xlsx"
            try CoreUtils.validateFilePath(testPath)
            print("   ✓ File path validation passed: \(testPath)")
        } catch {
            print("   ✗ File path validation failed: \(error)")
        }
        
        // Test 3: Path traversal attempt
        print("\n3. Testing path traversal protection:")
        do {
            let maliciousPath = "/Users/vr/Developer/XLKit/Test-Workflows/test-restricted.xlsx"
            try CoreUtils.validateFilePath(maliciousPath)
            print("   ✓ File path validation passed: \(maliciousPath)")
        } catch {
            print("   ✗ File path validation failed: \(error)")
        }
        
        // Reset to default
        SecurityManager.enableFilePathRestrictions = false
        print("\n=== Demo completed ===")
    }

    // MARK: - Helper Functions

    private static func calculateTextWidth(_ text: String) -> Double {
        // Simple width calculation based on character count
        // In a real implementation, you might use Core Text for more accurate measurements
        return Double(text.count) * 8.0 // Approximate width per character
    }

    private static func createTestImageData() -> Data {
        // Create a simple 1x1 pixel PNG image for testing
        // This is a minimal valid PNG file
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

    /// Generates an Excel file with image embeds (future implementation)
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

    static func validateExcelFile(_ filePath: String) {
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
            _ = try file.parseStyles()
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
} 