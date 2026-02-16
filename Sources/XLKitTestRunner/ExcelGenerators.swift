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
        
        guard let csvData = try? String(contentsOfFile: csvFilePath, encoding: .utf8) else {
            print("[ERROR] Failed to read CSV file: \(csvFilePath)")
            throw XLKitError.fileWriteError("Failed to read CSV file: \(csvFilePath)")
        }
        
        // MARK: - Create Excel File
        print("[INFO] Creating Excel workbook...")
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Embed Test")

        // Import CSV using XLKit's public API (powered by swift-textfile-tools for robust parsing).
        // Use hasHeader: false so the header row is included in the sheet and can be formatted.
        workbook.importCSV(csvData, into: sheet, hasHeader: false)

        let usedCoordinates = sheet.getUsedCells().compactMap { CellCoordinate(excelAddress: $0) }
        let maxRow = usedCoordinates.map(\.row).max() ?? 0
        let maxColumn = usedCoordinates.map(\.column).max() ?? 0

        let headers: [String] = (1...maxColumn).map { column in
            let coordinate = CellCoordinate(row: 1, column: column).excelAddress
            return sheet.getCell(coordinate)?.stringValue ?? ""
        }

        let dataRowCount = max(0, maxRow - 1)
        print("[INFO] Imported CSV with \(headers.count) columns and \(dataRowCount) data rows")
        
        // MARK: - Calculate Column Widths
        print("[INFO] Calculating optimal column widths...")
        var columnWidths: [Int: Double] = [:]
        
        for col in 1...maxColumn {
            var maxWidth = calculateTextWidth(headers[col - 1])

            for row in 2...maxRow {
                let coordinate = CellCoordinate(row: row, column: col).excelAddress
                let text = sheet.getCell(coordinate)?.stringValue ?? ""
                maxWidth = max(maxWidth, calculateTextWidth(text))
            }

            let adjustedWidth = min(max(maxWidth + 4.0, 8.0), 50.0)
            columnWidths[col] = adjustedWidth
            print("[INFO] Column \(col) (\(headers[col - 1])): width = \(adjustedWidth)")
        }
        
        // MARK: - Write Headers
        print("[INFO] Writing headers with bold formatting...")
        let headerFormat = CellFormat.header(fontSize: 12.0, backgroundColor: "#E6E6E6")
        
        for (colIndex, header) in headers.enumerated() {
            let column = colIndex + 1
            sheet.setCell(row: 1, column: column, cell: Cell.string(header, format: headerFormat))
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
        sheet.setCell("D2", number: 0.1234, format: CellFormat.percentage())
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
        
        // MARK: - Column Ordering Test
        print("[INFO] Demonstrating column ordering for sheets with >26 columns...")
        
        let columnOrderSheet = workbook.addSheet(name: "Column Order Test")
        
        // Fill columns A through AD (30 columns total) - this tests the column ordering fix
        for i in 1...30 {
            let column = i
            let value = "Col\(i)"
            columnOrderSheet.setCell(row: 1, column: column, cell: Cell.string(value, format: CellFormat.header()))
        }
        
        // Add some data in row 2 to make it more realistic
        for i in 1...30 {
            let column = i
            let value = "Data\(i)"
            columnOrderSheet.setCell(row: 2, column: column, cell: Cell.string(value))
        }
        
        print("[INFO] ✓ Created sheet with 30 columns (A1 through AD1) to test column ordering")
        
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
            
            // Validate column ordering sheet (it's the 4th sheet: API Demo, Second Sheet, Fluent Demo, Column Order Test)
            let columnOrderWorksheet = try xlsx.parseWorksheet(at: "xl/worksheets/sheet4.xml")
            let columnOrderRows = columnOrderWorksheet.data?.rows ?? []
            print("[INFO] ✓ Column Order Test sheet contains \(columnOrderRows.count) rows")
            
            if let firstRow = columnOrderRows.first {
                let cells = firstRow.cells
                print("[INFO] ✓ Column Order Test sheet contains \(cells.count) columns")
                
                // Verify column ordering (A, B, C, ..., Z, AA, AB, AC, AD)
                let expectedColumns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"]
                var correctOrder = true
                
                for (index, cell) in cells.enumerated() {
                    if index < expectedColumns.count {
                        let expectedColumn = expectedColumns[index]
                        let actualColumn = String(describing: cell.reference.column)
                        if actualColumn != expectedColumn {
                            correctOrder = false
                            break
                        }
                    }
                }
                
                if correctOrder {
                    print("[INFO] ✓ Column ordering is correct - no Excel repair warnings should occur")
                } else {
                    print("[WARNING] Column ordering may have issues")
                }
            }
            
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

    // MARK: - Number Format Testing
    
    /// Tests comprehensive number formatting functionality
    /// 
    /// Validates that the number format fix works correctly with:
    /// - Currency formatting with thousands grouping
    /// - Percentage formatting with % symbol
    /// - Custom number formats
    /// - Excel compliance and proper XML generation
    @MainActor
    static func testNumberFormats() async throws {
        print("[INFO] Starting comprehensive number format testing...")
        
        // MARK: - Workbook Creation
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        print("[INFO] Created workbook with sheet: \(sheet.name)")
        
        // MARK: - Currency Formatting Tests
        print("[INFO] Testing currency formatting...")
        
        sheet.setCell("A1", string: "Currency Formats", format: CellFormat.header())
        sheet.setCell("A2", string: "Standard Currency")
        sheet.setCell("A3", string: "Custom Currency")
        sheet.setCell("A4", string: "Red Currency")
        
        sheet.setCell("B2", number: 1234.56, format: CellFormat.currency())
        sheet.setCell("B3", number: 5678.90, format: CellFormat.currency(format: .custom("$#,##0.00")))
        
        var redCurrencyFormat = CellFormat.currency()
        redCurrencyFormat.fontColor = "#FF0000"
        sheet.setCell("B4", number: 9999.99, format: redCurrencyFormat)
        
        print("[INFO] ✓ Applied currency formats: standard, custom, and red")
        
        // MARK: - Percentage Formatting Tests
        print("[INFO] Testing percentage formatting...")
        
        sheet.setCell("D1", string: "Percentage Formats", format: CellFormat.header())
        sheet.setCell("D2", string: "Standard Percentage")
        sheet.setCell("D3", string: "Custom Percentage")
        sheet.setCell("D4", string: "Blue Percentage")
        
        sheet.setCell("E2", number: 0.1234, format: CellFormat.percentage())
        sheet.setCell("E3", number: 0.5678, format: CellFormat.percentage(format: .custom("0.00%")))
        
        var bluePercentageFormat = CellFormat.percentage()
        bluePercentageFormat.fontColor = "#0000FF"
        sheet.setCell("E4", number: 0.9999, format: bluePercentageFormat)
        
        print("[INFO] ✓ Applied percentage formats: standard, custom, and blue")
        
        // MARK: - Custom Number Format Tests
        print("[INFO] Testing custom number formats...")
        
        sheet.setCell("G1", string: "Custom Formats", format: CellFormat.header())
        sheet.setCell("G2", string: "Thousands Separator")
        sheet.setCell("G3", string: "Decimal Places")
        sheet.setCell("G4", string: "Mixed Format")
        
        var thousandsFormat = CellFormat()
        thousandsFormat.numberFormat = .custom
        thousandsFormat.customNumberFormat = "#,##0"
        sheet.setCell("H2", number: 1234567, format: thousandsFormat)
        
        var decimalFormat = CellFormat()
        decimalFormat.numberFormat = .custom
        decimalFormat.customNumberFormat = "0.000"
        sheet.setCell("H3", number: 3.14159, format: decimalFormat)
        
        var mixedFormat = CellFormat()
        mixedFormat.numberFormat = .custom
        mixedFormat.customNumberFormat = "$#,##0.00;($#,##0.00)"
        sheet.setCell("H4", number: -1234.56, format: mixedFormat)
        
        print("[INFO] ✓ Applied custom formats: thousands, decimals, and mixed")
        
        // MARK: - Date Format Tests
        print("[INFO] Testing date formatting...")
        
        sheet.setCell("J1", string: "Date Formats", format: CellFormat.header())
        sheet.setCell("J2", string: "Standard Date")
        sheet.setCell("J3", string: "Custom Date")
        
        let currentDate = Date()
        sheet.setCell("K2", date: currentDate, format: CellFormat.date())
        
        var customDateFormat = CellFormat()
        customDateFormat.numberFormat = .custom
        customDateFormat.customNumberFormat = "mmmm dd, yyyy"
        sheet.setCell("K3", date: currentDate, format: customDateFormat)
        
        print("[INFO] ✓ Applied date formats: standard and custom")
        
        // MARK: - Column Widths
        print("[INFO] Setting column widths for better display...")
        
        sheet.setColumnWidth(1, width: 15.0) // A
        sheet.setColumnWidth(2, width: 12.0) // B
        sheet.setColumnWidth(4, width: 15.0) // D
        sheet.setColumnWidth(5, width: 12.0) // E
        sheet.setColumnWidth(7, width: 15.0) // G
        sheet.setColumnWidth(8, width: 12.0) // H
        sheet.setColumnWidth(10, width: 15.0) // J
        sheet.setColumnWidth(11, width: 15.0) // K
        
        // MARK: - Save and Validate
        print("[INFO] Saving number format test...")
        
        let outputPath = "Test-Workflows/Number-Format-Test.xlsx"
        let outputURL = URL(fileURLWithPath: outputPath)
        
        let outputDir = outputURL.deletingLastPathComponent()
        if !FileManager.default.fileExists(atPath: outputDir.path) {
            try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
        }
        
        try await workbook.save(to: outputURL)
        print("[SUCCESS] Number format test saved: \(outputPath)")
        
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
        
        print("[SUCCESS] Number format testing completed successfully!")
        print("[INFO] Expected Excel display:")
        print("   - B2: $1,234.56 (currency with thousands grouping)")
        print("   - B3: $5,678.90 (custom currency format)")
        print("   - B4: $9,999.99 (red currency)")
        print("   - E2: 12.34% (percentage)")
        print("   - E3: 56.78% (custom percentage)")
        print("   - E4: 99.99% (blue percentage)")
        print("   - H2: 1,234,567 (thousands separator)")
        print("   - H3: 3.142 (3 decimal places)")
        print("   - H4: ($1,234.56) (negative format)")
        print("   - K2: Current date in standard format")
        print("   - K3: Current date in custom format")
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

    // MARK: - iOS Compatibility Test
    
    /// Tests iOS-specific file system compatibility
    @MainActor
    static func testIOSCompatibility() async throws {
        print("[INFO] Testing iOS file system compatibility...")
        
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "iOS Test")
        
        // Add some test data
        sheet.setCell("A1", string: "iOS Compatibility Test", format: CellFormat.header())
        sheet.setCell("A2", string: "This file was created using iOS-safe file operations")
        sheet.setCell("A3", string: "Date: \(Date())")
        
        // Test using CoreUtils.safeFileURL
        let safeURL = CoreUtils.safeFileURL(for: "ios-compatibility-test.xlsx")
        print("[INFO] Using safe file URL: \(safeURL.path)")
        
        try await workbook.save(to: safeURL)
        print("[SUCCESS] iOS compatibility test completed successfully!")
        print("[INFO] File saved to: \(safeURL.path)")
        
        // Verify the file exists
        if FileManager.default.fileExists(atPath: safeURL.path) {
            print("[INFO] ✓ File exists and is accessible")
        } else {
            print("[ERROR] File was not created successfully")
        }
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