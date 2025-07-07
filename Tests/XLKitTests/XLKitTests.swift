//
//  XLKitTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

final class XLKitTests: XCTestCase {
    
    func testCreateWorkbook() {
        let workbook = XLKit.createWorkbook()
        XCTAssertNotNil(workbook)
        XCTAssertEqual(workbook.getSheets().count, 0)
    }
    
    func testAddSheet() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test Sheet")
        
        XCTAssertEqual(workbook.getSheets().count, 1)
        XCTAssertEqual(sheet.name, "Test Sheet")
        XCTAssertEqual(sheet.id, 1)
    }
    
    func testGetSheetByName() {
        let workbook = XLKit.createWorkbook()
        let sheet1 = workbook.addSheet(name: "Sheet1")
        let sheet2 = workbook.addSheet(name: "Sheet2")
        
        XCTAssertEqual(workbook.getSheet(name: "Sheet1"), sheet1)
        XCTAssertEqual(workbook.getSheet(name: "Sheet2"), sheet2)
        XCTAssertNil(workbook.getSheet(name: "NonExistent"))
    }
    
    func testRemoveSheet() {
        let workbook = XLKit.createWorkbook()
        _ = workbook.addSheet(name: "Sheet1")
        _ = workbook.addSheet(name: "Sheet2")
        
        XCTAssertEqual(workbook.getSheets().count, 2)
        
        workbook.removeSheet(name: "Sheet1")
        XCTAssertEqual(workbook.getSheets().count, 1)
        XCTAssertNil(workbook.getSheet(name: "Sheet1"))
        XCTAssertNotNil(workbook.getSheet(name: "Sheet2"))
    }
    
    func testSetAndGetCell() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test string value
        sheet.setCell("A1", value: .string("Hello"))
        XCTAssertEqual(sheet.getCell("A1"), .string("Hello"))
        
        // Test number value
        sheet.setCell("B1", value: .number(42.5))
        XCTAssertEqual(sheet.getCell("B1"), .number(42.5))
        
        // Test integer value
        sheet.setCell("C1", value: .integer(100))
        XCTAssertEqual(sheet.getCell("C1"), .integer(100))
        
        // Test boolean value
        sheet.setCell("D1", value: .boolean(true))
        XCTAssertEqual(sheet.getCell("D1"), .boolean(true))
        
        // Test date value
        let date = Date()
        sheet.setCell("E1", value: .date(date))
        XCTAssertEqual(sheet.getCell("E1"), .date(date))
        
        // Test formula
        sheet.setCell("F1", value: .formula("=A1+B1"))
        XCTAssertEqual(sheet.getCell("F1"), .formula("=A1+B1"))
    }
    
    func testSetCellByRowColumn() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell(row: 1, column: 1, value: .string("A1"))
        sheet.setCell(row: 2, column: 2, value: .string("B2"))
        
        XCTAssertEqual(sheet.getCell("A1"), .string("A1"))
        XCTAssertEqual(sheet.getCell("B2"), .string("B2"))
    }
    
    func testSetRange() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRange("A1:C3", string: "Range")
        
        XCTAssertEqual(sheet.getCell("A1"), .string("Range"))
        XCTAssertEqual(sheet.getCell("B2"), .string("Range"))
        XCTAssertEqual(sheet.getCell("C3"), .string("Range"))
    }
    
    func testMergeCells() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.mergeCells("A1:C1")
        
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, 1)
        XCTAssertEqual(mergedRanges[0].excelRange, "A1:C1")
    }
    
    func testGetUsedCells() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell("A1", value: .string("A1"))
        sheet.setCell("B2", value: .string("B2"))
        sheet.setCell("C3", value: .string("C3"))
        
        let usedCells = sheet.getUsedCells()
        XCTAssertEqual(usedCells.count, 3)
        XCTAssertTrue(usedCells.contains("A1"))
        XCTAssertTrue(usedCells.contains("B2"))
        XCTAssertTrue(usedCells.contains("C3"))
    }
    
    func testClearSheet() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell("A1", value: .string("Test"))
        sheet.mergeCells("B1:C1")
        
        XCTAssertEqual(sheet.getUsedCells().count, 1)
        XCTAssertEqual(sheet.getMergedRanges().count, 1)
        
        sheet.clear()
        
        XCTAssertEqual(sheet.getUsedCells().count, 0)
        XCTAssertEqual(sheet.getMergedRanges().count, 0)
    }
    
    func testCellCoordinate() {
        // Test coordinate creation
        let coord1 = CellCoordinate(row: 1, column: 1)
        XCTAssertEqual(coord1.excelAddress, "A1")
        
        let coord2 = CellCoordinate(row: 2, column: 2)
        XCTAssertEqual(coord2.excelAddress, "B2")
        
        // Test coordinate from Excel address
        let coord3 = CellCoordinate(excelAddress: "A1")
        XCTAssertNotNil(coord3)
        XCTAssertEqual(coord3?.row, 1)
        XCTAssertEqual(coord3?.column, 1)
        
        let coord4 = CellCoordinate(excelAddress: "B2")
        XCTAssertNotNil(coord4)
        XCTAssertEqual(coord4?.row, 2)
        XCTAssertEqual(coord4?.column, 2)
        
        // Test invalid address
        XCTAssertNil(CellCoordinate(excelAddress: "Invalid"))
    }
    
    func testCellRange() {
        // Test range creation
        let start = CellCoordinate(row: 1, column: 1)
        let end = CellCoordinate(row: 3, column: 3)
        let range = CellRange(start: start, end: end)
        
        XCTAssertEqual(range.excelRange, "A1:C3")
        XCTAssertEqual(range.coordinates.count, 9) // 3x3 grid
        
        // Test range from Excel notation
        let range2 = CellRange(excelRange: "A1:B2")
        XCTAssertNotNil(range2)
        XCTAssertEqual(range2?.excelRange, "A1:B2")
        XCTAssertEqual(range2?.coordinates.count, 4) // 2x2 grid
        
        // Test invalid range
        XCTAssertNil(CellRange(excelRange: "Invalid"))
    }
    
    func testCellValueStringValue() {
        XCTAssertEqual(CellValue.string("Hello").stringValue, "Hello")
        XCTAssertEqual(CellValue.number(42.5).stringValue, "42.5")
        XCTAssertEqual(CellValue.integer(100).stringValue, "100")
        XCTAssertEqual(CellValue.boolean(true).stringValue, "TRUE")
        XCTAssertEqual(CellValue.boolean(false).stringValue, "FALSE")
        XCTAssertEqual(CellValue.formula("=A1+B1").stringValue, "=A1+B1")
        XCTAssertEqual(CellValue.empty.stringValue, "")
    }
    
    func testCellValueType() {
        XCTAssertEqual(CellValue.string("").type, "string")
        XCTAssertEqual(CellValue.number(0).type, "number")
        XCTAssertEqual(CellValue.integer(0).type, "integer")
        XCTAssertEqual(CellValue.boolean(true).type, "boolean")
        XCTAssertEqual(CellValue.date(Date()).type, "date")
        XCTAssertEqual(CellValue.formula("").type, "formula")
        XCTAssertEqual(CellValue.empty.type, "empty")
    }
    
    func testColumnConversion() {
        XCTAssertEqual(CoreUtils.columnLetter(from: 1), "A")
        XCTAssertEqual(CoreUtils.columnLetter(from: 2), "B")
        XCTAssertEqual(CoreUtils.columnLetter(from: 26), "Z")
        XCTAssertEqual(CoreUtils.columnLetter(from: 27), "AA")
        XCTAssertEqual(CoreUtils.columnLetter(from: 28), "AB")
        
        XCTAssertEqual(CoreUtils.columnNumber(from: "A"), 1)
        XCTAssertEqual(CoreUtils.columnNumber(from: "B"), 2)
        XCTAssertEqual(CoreUtils.columnNumber(from: "Z"), 26)
        XCTAssertEqual(CoreUtils.columnNumber(from: "AA"), 27)
        XCTAssertEqual(CoreUtils.columnNumber(from: "AB"), 28)
    }
    
    func testDateConversion() {
        let date = Date()
        let excelNumber = CoreUtils.excelNumberFromDate(date)
        let convertedDate = CoreUtils.dateFromExcelNumber(excelNumber)
        
        // Allow for small precision differences
        XCTAssertEqual(date.timeIntervalSince1970, convertedDate.timeIntervalSince1970, accuracy: 1.0)
    }
    
    func testXMLEscaping() {
        XCTAssertEqual(CoreUtils.escapeXML("&"), "&amp;")
        XCTAssertEqual(CoreUtils.escapeXML("<"), "&lt;")
        XCTAssertEqual(CoreUtils.escapeXML(">"), "&gt;")
        XCTAssertEqual(CoreUtils.escapeXML("\""), "&quot;")
        XCTAssertEqual(CoreUtils.escapeXML("'"), "&apos;")
        XCTAssertEqual(CoreUtils.escapeXML("Hello & World"), "Hello &amp; World")
    }
    
    func testSaveWorkbook() async throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell("A1", value: .string("Hello"))
        sheet.setCell("B1", value: .number(42))
        sheet.setCell("C1", value: .boolean(true))
        
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("test.xlsx")
        defer {
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
    }
    
    func testSaveWorkbookSync() throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell("A1", value: .string("Hello"))
        sheet.setCell("B1", value: .number(42))
        
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("test_sync.xlsx")
        defer {
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try XLKit.saveWorkbook(workbook, to: tempURL)
        
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
    }
    
    func testComplexWorkbook() async throws {
        let workbook = XLKit.createWorkbook()
        
        // Add multiple sheets
        let sheet1 = workbook.addSheet(name: "Data")
        let sheet2 = workbook.addSheet(name: "Summary")
        
        // Add data to first sheet
        sheet1.setCell("A1", value: .string("Name"))
        sheet1.setCell("B1", value: .string("Age"))
        sheet1.setCell("C1", value: .string("Salary"))
        
        sheet1.setCell("A2", value: .string("John"))
        sheet1.setCell("B2", value: .integer(30))
        sheet1.setCell("C2", value: .number(50000.0))
        
        sheet1.setCell("A3", value: .string("Jane"))
        sheet1.setCell("B3", value: .integer(25))
        sheet1.setCell("C3", value: .number(45000.0))
        
        // Add summary to second sheet
        sheet2.setCell("A1", value: .string("Total Employees"))
        sheet2.setCell("B1", value: .formula("=COUNTA(Data!A:A)-1"))
        
        sheet2.setCell("A2", value: .string("Average Age"))
        sheet2.setCell("B2", value: .formula("=AVERAGE(Data!B:B)"))
        
        sheet2.setCell("A3", value: .string("Total Salary"))
        sheet2.setCell("B3", value: .formula("=SUM(Data!C:C)"))
        
        // Merge cells for headers
        sheet1.mergeCells("A1:C1")
        sheet2.mergeCells("A1:B1")
        
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("complex.xlsx")
        defer {
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
    }
    
    // MARK: - Image Support Tests
    
    func testImageFormatDetection() {
        // Test GIF format detection
        let gifData = Data([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x10, 0x00])
        XCTAssertEqual(ImageUtils.detectImageFormat(from: gifData), .gif)
        
        // Test PNG format detection
        let pngData = Data([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])
        XCTAssertEqual(ImageUtils.detectImageFormat(from: pngData), .png)
        
        // Test JPEG format detection
        let jpegData = Data([0xFF, 0xD8, 0xFF, 0xE0])
        XCTAssertEqual(ImageUtils.detectImageFormat(from: jpegData), .jpeg)
        
        // Test invalid data
        let invalidData = Data([0x00, 0x00, 0x00, 0x00])
        XCTAssertNil(ImageUtils.detectImageFormat(from: invalidData))
    }
    
    func testImageSizeDetection() {
        // Test GIF size detection
        let gifData = Data([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00])
        let gifSize = ImageUtils.getImageSize(from: gifData, format: .gif)
        XCTAssertEqual(gifSize?.width, 16)
        XCTAssertEqual(gifSize?.height, 32)
        
        // Test PNG size detection
        let pngData = Data([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x00, 0x30])
        let pngSize = ImageUtils.getImageSize(from: pngData, format: .png)
        XCTAssertEqual(pngSize?.width, 64)
        XCTAssertEqual(pngSize?.height, 48)
    }
    
    func testExcelImageCreation() {
        let imageData = Data([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00])
        
        // Test creation from data
        let image = ImageUtils.createExcelImage(from: imageData, format: .gif)
        XCTAssertNotNil(image)
        XCTAssertEqual(image?.format, .gif)
        XCTAssertEqual(image?.originalSize.width, 16)
        XCTAssertEqual(image?.originalSize.height, 32)
        
        // Test creation with display size
        let imageWithSize = ImageUtils.createExcelImage(from: imageData, format: .gif, displaySize: CGSize(width: 100, height: 200))
        XCTAssertNotNil(imageWithSize)
        XCTAssertEqual(imageWithSize?.displaySize?.width, 100)
        XCTAssertEqual(imageWithSize?.displaySize?.height, 200)
    }
    
    func testSheetImageOperations() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Images")
        
        let imageData = Data([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00])
        
        // Test adding image from data
        let success = sheet.addImage(imageData, at: "A1", format: .gif)
        XCTAssertTrue(success)
        
        // Test getting images
        let images = sheet.getImages()
        XCTAssertEqual(images.count, 1)
        XCTAssertNotNil(images["A1"])
        XCTAssertEqual(images["A1"]?.format, .gif)
    }
    
    func testWorkbookImageOperations() {
        let workbook = XLKit.createWorkbook()
        
        let imageData = Data([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00])
        let image = ImageUtils.createExcelImage(from: imageData, format: .gif)!
        
        // Test adding image to workbook
        workbook.addImage(image)
        
        // Test getting images
        let images = workbook.getImages()
        XCTAssertEqual(images.count, 1)
        XCTAssertEqual(images[0].format, .gif)
    }
    
    // MARK: - Column and Row Sizing Tests
    
    func testColumnWidthOperations() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Sizing")
        
        // Test setting column width
        sheet.setColumnWidth(1, width: 150.0)
        sheet.setColumnWidth(2, width: 200.0)
        
        // Test getting column width
        XCTAssertEqual(sheet.getColumnWidth(1), 150.0)
        XCTAssertEqual(sheet.getColumnWidth(2), 200.0)
        XCTAssertNil(sheet.getColumnWidth(3))
        
        // Test getting all column widths
        let widths = sheet.getColumnWidths()
        XCTAssertEqual(widths.count, 2)
        XCTAssertEqual(widths[1], 150.0)
        XCTAssertEqual(widths[2], 200.0)
    }
    
    func testRowHeightOperations() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Sizing")
        
        // Test setting row height
        sheet.setRowHeight(1, height: 25.0)
        sheet.setRowHeight(2, height: 30.0)
        
        // Test getting row height
        XCTAssertEqual(sheet.getRowHeight(1), 25.0)
        XCTAssertEqual(sheet.getRowHeight(2), 30.0)
        XCTAssertNil(sheet.getRowHeight(3))
        
        // Test getting all row heights
        let heights = sheet.getRowHeights()
        XCTAssertEqual(heights.count, 2)
        XCTAssertEqual(heights[1], 25.0)
        XCTAssertEqual(heights[2], 30.0)
    }
    
    func testAutoSizeColumnForImage() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "AutoSize")
        
        let imageData = Data([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00])
        let success = sheet.addImage(imageData, at: "A1", format: .gif, displaySize: CGSize(width: 150, height: 100))
        XCTAssertTrue(success)
        
        // Test auto-sizing column
        sheet.autoSizeColumn(1, forImageAt: "A1")
        XCTAssertEqual(sheet.getColumnWidth(1), 150.0)
    }
    
    // MARK: - GIF Support Tests
    
    func testGIFEmbedding() async throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "GIF Test")
        
        // Create a simple GIF data (minimal valid GIF)
        let gifData = createMinimalGIF()
        
        // Add GIF to sheet
        let success = sheet.addImage(gifData, at: "A1", format: .gif)
        XCTAssertTrue(success)
        
        // Set column width to fit GIF
        sheet.autoSizeColumn(1, forImageAt: "A1")
        
        // Save workbook
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("gif_test.xlsx")
        defer {
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
    }
    
    func testMultipleImageFormats() async throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Multiple Formats")
        
        // Add different image formats
        let gifData = createMinimalGIF()
        let pngData = createMinimalPNG()
        
        sheet.addImage(gifData, at: "A1", format: .gif)
        sheet.addImage(pngData, at: "B1", format: .png)
        
        // Set column widths
        sheet.autoSizeColumn(1, forImageAt: "A1")
        sheet.autoSizeColumn(2, forImageAt: "B1")
        
        // Save workbook
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("multiple_formats.xlsx")
        defer {
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
    }
    
    // MARK: - Helper Methods
    
    private func createMinimalGIF() -> Data {
        // Minimal valid GIF data (1x1 transparent pixel)
        let gifBytes: [UInt8] = [
            0x47, 0x49, 0x46, 0x38, 0x39, 0x61, // GIF89a header
            0x01, 0x00, // Width: 1
            0x01, 0x00, // Height: 1
            0x91, // Color table info
            0x00, // Background color
            0x00, // Aspect ratio
            0x00, 0x00, 0x00, // Color table (empty)
            0x21, 0xF9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00, // Graphics control extension
            0x2C, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, // Image descriptor
            0x02, 0x02, 0x44, 0x01, 0x00, // Image data
            0x3B // Trailer
        ]
        return Data(gifBytes)
    }
    
    private func createMinimalPNG() -> Data {
        // Minimal valid PNG data (1x1 transparent pixel)
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0x01, // Width: 1
            0x00, 0x00, 0x00, 0x01, // Height: 1
            0x08, 0x06, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x1F, 0x15, 0xC4, 0x89, // IHDR CRC
            0x00, 0x00, 0x00, 0x0C, // IDAT chunk length
            0x49, 0x44, 0x41, 0x54, // IDAT
            0x08, 0x99, 0x01, 0x01, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, // IDAT data
            0xE2, 0x21, 0xBC, 0x33, // IDAT CRC
            0x00, 0x00, 0x00, 0x00, // IEND chunk length
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // IEND CRC
        ]
        return Data(pngBytes)
    }
    
    private func createWidePNG() -> Data {
        // Create a minimal valid PNG with 16:9 aspect ratio (160x90)
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0xA0, // Width: 160 (big endian)
            0x00, 0x00, 0x00, 0x5A, // Height: 90 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    private func createTallPNG() -> Data {
        // Create a minimal valid PNG with 9:16 aspect ratio (90x160)
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0x5A, // Width: 90 (big endian)
            0x00, 0x00, 0x00, 0xA0, // Height: 160 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    private func createSquarePNG() -> Data {
        // Create a minimal valid PNG with 1:1 aspect ratio (100x100)
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0x64, // Width: 100 (big endian)
            0x00, 0x00, 0x00, 0x64, // Height: 100 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    private func createUltraWidePNG() -> Data {
        // Create a minimal valid PNG with 21:9 aspect ratio (210x90)
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0xD2, // Width: 210 (big endian)
            0x00, 0x00, 0x00, 0x5A, // Height: 90 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    private func createPortraitPNG() -> Data {
        // Create a minimal valid PNG with 3:4 aspect ratio (75x100)
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0x4B, // Width: 75 (big endian)
            0x00, 0x00, 0x00, 0x64, // Height: 100 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    // MARK: - Movie/Video Aspect Ratio PNGs
    
    private func createCinemascopePNG() -> Data {
        // Create a minimal valid PNG with 2.39:1 aspect ratio (239x100) - Cinemascope/Anamorphic
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0xEF, // Width: 239 (big endian)
            0x00, 0x00, 0x00, 0x64, // Height: 100 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    private func createAcademyPNG() -> Data {
        // Create a minimal valid PNG with 1.85:1 aspect ratio (185x100) - Academy ratio
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0xB9, // Width: 185 (big endian)
            0x00, 0x00, 0x00, 0x64, // Height: 100 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    private func createClassicTVPNG() -> Data {
        // Create a minimal valid PNG with 4:3 aspect ratio (133x100) - Classic TV/monitor
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0x85, // Width: 133 (big endian)
            0x00, 0x00, 0x00, 0x64, // Height: 100 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    private func createModernMobilePNG() -> Data {
        // Create a minimal valid PNG with 18:9 aspect ratio (180x90) - Modern smartphone screens
        let pngBytes: [UInt8] = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, // IHDR chunk length
            0x49, 0x48, 0x44, 0x52, // IHDR
            0x00, 0x00, 0x00, 0xB4, // Width: 180 (big endian)
            0x00, 0x00, 0x00, 0x5A, // Height: 90 (big endian)
            0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, compression, filter, interlace
            0x00, 0x00, 0x00, 0x00, // CRC placeholder
            0x00, 0x00, 0x00, 0x00, // IEND chunk
            0x49, 0x45, 0x4E, 0x44, // IEND
            0xAE, 0x42, 0x60, 0x82  // CRC
        ]
        return Data(pngBytes)
    }
    
    // MARK: - Simplified Image Embedding API Tests
    
    func testSimplifiedImageEmbeddingAPI() async throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Simplified API Test")
        
        // Create a simple PNG data (minimal valid PNG)
        let pngData = createMinimalPNG()
        
        // Test the simplified embedImage API
        let success = sheet.embedImage(
            pngData,
            at: "A1",
            of: workbook,
            scale: 0.8, // 80% of max size
            maxWidth: 400,
            maxHeight: 300
        )
        
        XCTAssertTrue(success)
        
        // Verify image was added
        let images = sheet.getImages()
        XCTAssertEqual(images.count, 1)
        XCTAssertNotNil(images["A1"])
        
        // Verify workbook has the image
        XCTAssertEqual(workbook.imageCount, 1)
        
        // Save workbook to test file generation
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("simplified_api_test.xlsx")
        defer {
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
    }
    
    func testAspectRatioPreservation() async throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Aspect Ratio Test")
        
        // Create a wide rectangular PNG (16:9 aspect ratio)
        let widePNGData = createWidePNG()
        
        // Embed with the simplified API
        let success = sheet.embedImage(
            widePNGData,
            at: "A1",
            of: workbook,
            scale: 0.5, // 50% of max size
            maxWidth: 600,
            maxHeight: 400
        )
        
        XCTAssertTrue(success)
        
        // Get the embedded image
        let images = sheet.getImages()
        guard let image = images["A1"] else {
            XCTFail("Image not found")
            return
        }
        
        // Verify aspect ratio is preserved
        let originalRatio = image.originalSize.width / image.originalSize.height
        let displayRatio = image.displaySize!.width / image.displaySize!.height
        
        // Aspect ratios should be exactly equal (within floating point precision)
        XCTAssertEqual(originalRatio, displayRatio, accuracy: 0.001)
        
        // Verify Excel cell dimensions match the display size
        let cellCoord = CellCoordinate(excelAddress: "A1")!
        let colWidth = sheet.getColumnWidth(cellCoord.column)
        let rowHeight = sheet.getRowHeight(cellCoord.row)
        
        // Calculate expected Excel dimensions from display size
        let expectedColWidth = ImageSizingUtils.excelColumnWidth(forPixelWidth: image.displaySize!.width)
        let expectedRowHeight = ImageSizingUtils.excelRowHeight(forPixelHeight: image.displaySize!.height)
        
        XCTAssertEqual(colWidth ?? 0, CGFloat(expectedColWidth), accuracy: 0.001)
        XCTAssertEqual(rowHeight ?? 0, CGFloat(expectedRowHeight), accuracy: 0.001)
        
        // Save workbook
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("aspect_ratio_test.xlsx")
        defer {
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
    }
    
    func testAspectRatioPreservationWithDifferentSizes() async throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Aspect Ratio Test Different Sizes")
        
        // Test with different aspect ratios
        let testCases = [
            ("A1", createWidePNG(), "16:9 wide"),
            ("B1", createSquarePNG(), "1:1 square"),
            ("C1", createTallPNG(), "9:16 tall")
        ]
        
        for (coordinate, imageData, description) in testCases {
            let success = sheet.embedImage(
                imageData,
                at: coordinate,
                of: workbook,
                scale: 0.8,
                maxWidth: 500,
                maxHeight: 300
            )
            
            XCTAssertTrue(success, "Failed to embed \(description)")
            
            // Get the embedded image
            let images = sheet.getImages()
            guard let image = images[coordinate] else {
                XCTFail("Image not found for \(description)")
                continue
            }
            
            // Verify aspect ratio is preserved
            let originalRatio = image.originalSize.width / image.originalSize.height
            let displayRatio = image.displaySize!.width / image.displaySize!.height
            
            XCTAssertEqual(originalRatio, displayRatio, accuracy: 0.001, "Aspect ratio not preserved for \(description)")
            
            // Verify the same scale was applied to both dimensions
            let widthScale = image.displaySize!.width / image.originalSize.width
            let heightScale = image.displaySize!.height / image.originalSize.height
            XCTAssertEqual(widthScale, heightScale, accuracy: 0.001, "Different scales applied for \(description)")
        }
        
        // Save workbook
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("aspect_ratio_different_sizes_test.xlsx")
        defer {
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
    }
    
    func testAspectRatioPreservationForAnyDimensions() async throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Any Dimensions Test")
        
        // Test with various image dimensions and aspect ratios
        let testCases = [
            (createWidePNG(), "16:9 wide (160x90)"),
            (createSquarePNG(), "1:1 square (100x100)"),
            (createTallPNG(), "9:16 tall (90x160)"),
            (createUltraWidePNG(), "21:9 ultra-wide (210x90)"),
            (createPortraitPNG(), "3:4 portrait (75x100)"),
            (createCinemascopePNG(), "2.39:1 cinemascope (239x100)"),
            (createAcademyPNG(), "1.85:1 academy (185x100)"),
            (createClassicTVPNG(), "4:3 classic TV (133x100)"),
            (createModernMobilePNG(), "18:9 modern mobile (180x90)")
        ]
        
        for (index, (imageData, description)) in testCases.enumerated() {
            let coordinate = "A\(index + 1)"
            
            // Test the simplified embedImage API
            let success = sheet.embedImage(
                imageData,
                at: coordinate,
                of: workbook,
                scale: 0.6, // 60% of max size
                maxWidth: 500,
                maxHeight: 400
            )
            
            XCTAssertTrue(success, "Failed to embed \(description)")
            
            // Get the embedded image
            let images = sheet.getImages()
            guard let image = images[coordinate] else {
                XCTFail("Image not found for \(description)")
                continue
            }
            
            // Verify display size was calculated
            XCTAssertNotNil(image.displaySize, "Display size not set for \(description)")
            
            // Verify aspect ratio is preserved exactly
            let originalAspectRatio = image.originalSize.width / image.originalSize.height
            let displayAspectRatio = image.displaySize!.width / image.displaySize!.height
            XCTAssertEqual(originalAspectRatio, displayAspectRatio, accuracy: 0.001, 
                          "Aspect ratio not preserved for \(description)")
            
            print("[TEST] \(description): original=\(image.originalSize), display=\(image.displaySize!), aspect=\(originalAspectRatio)")
        }
        
        // Save and validate the workbook
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("AnyDimensionsTest.xlsx")
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        
        // Verify file was created
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
        
        // Clean up
        try FileManager.default.removeItem(at: tempURL)
    }
    
    // MARK: - Rich Cell Formatting Tests
    
    func testCellFormatting() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Formatting")
        
        // Test text formatting
        let headerFormat = CellFormat.header(fontSize: 16.0, backgroundColor: "#CCCCCC")
        sheet.setCell("A1", string: "Header", format: headerFormat)
        
        // Test currency formatting
        let currencyFormat = CellFormat.currency(format: .currencyWithDecimals, color: "#006600")
        sheet.setCell("B1", number: 1234.56, format: currencyFormat)
        
        // Test percentage formatting
        let percentageFormat = CellFormat.percentage()
        sheet.setCell("C1", number: 0.85, format: percentageFormat)
        
        // Test date formatting
        let dateFormat = CellFormat.date()
        let date = Date()
        sheet.setCell("D1", date: date, format: dateFormat)
        
        // Test bordered formatting
        let borderedFormat = CellFormat.bordered(style: .thin, color: "#000000")
        sheet.setCell("E1", string: "Bordered", format: borderedFormat)
        
        // Verify formatting is stored
        XCTAssertNotNil(sheet.getCellFormat("A1"))
        XCTAssertNotNil(sheet.getCellFormat("B1"))
        XCTAssertNotNil(sheet.getCellFormat("C1"))
        XCTAssertNotNil(sheet.getCellFormat("D1"))
        XCTAssertNotNil(sheet.getCellFormat("E1"))
        
        // Test getting cell with format
        let cellWithFormat = sheet.getCellWithFormat("A1")
        XCTAssertNotNil(cellWithFormat)
        XCTAssertEqual(cellWithFormat?.value, .string("Header"))
        XCTAssertNotNil(cellWithFormat?.format)
    }
    
    func testCustomCellFormatting() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Custom Formatting")
        
        // Create custom format
        var customFormat = CellFormat()
        customFormat.fontName = "Arial"
        customFormat.fontSize = 12.0
        customFormat.fontWeight = .bold
        customFormat.fontStyle = .italic
        customFormat.textDecoration = .underline
        customFormat.fontColor = "#FF0000"
        customFormat.backgroundColor = "#FFFF00"
        customFormat.horizontalAlignment = .center
        customFormat.verticalAlignment = .center
        customFormat.textWrapping = true
        customFormat.textRotation = 45
        customFormat.numberFormat = .custom("0.00%")
        customFormat.customNumberFormat = "0.00%"
        customFormat.borderTop = .thick
        customFormat.borderBottom = .thick
        customFormat.borderLeft = .thin
        customFormat.borderRight = .thin
        customFormat.borderColor = "#0000FF"
        
        sheet.setCell("A1", string: "Custom Formatted", format: customFormat)
        
        // Verify custom format is stored
        let retrievedFormat = sheet.getCellFormat("A1")
        XCTAssertNotNil(retrievedFormat)
        XCTAssertEqual(retrievedFormat?.fontName, "Arial")
        XCTAssertEqual(retrievedFormat?.fontSize, 12.0)
        XCTAssertEqual(retrievedFormat?.fontWeight, .bold)
        XCTAssertEqual(retrievedFormat?.fontStyle, .italic)
        XCTAssertEqual(retrievedFormat?.textDecoration, .underline)
        XCTAssertEqual(retrievedFormat?.fontColor, "#FF0000")
        XCTAssertEqual(retrievedFormat?.backgroundColor, "#FFFF00")
        XCTAssertEqual(retrievedFormat?.horizontalAlignment, .center)
        XCTAssertEqual(retrievedFormat?.verticalAlignment, .center)
        XCTAssertEqual(retrievedFormat?.textWrapping, true)
        XCTAssertEqual(retrievedFormat?.textRotation, 45)
        XCTAssertEqual(retrievedFormat?.numberFormat, .custom)
        XCTAssertEqual(retrievedFormat?.customNumberFormat, "0.00%")
        XCTAssertEqual(retrievedFormat?.borderTop, .thick)
        XCTAssertEqual(retrievedFormat?.borderBottom, .thick)
        XCTAssertEqual(retrievedFormat?.borderLeft, .thin)
        XCTAssertEqual(retrievedFormat?.borderRight, .thin)
        XCTAssertEqual(retrievedFormat?.borderColor, "#0000FF")
    }
    
    func testRangeFormatting() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Range Formatting")
        
        // Format a range of cells
        let headerFormat = CellFormat.header()
        sheet.setRange("A1:C1", string: "Header", format: headerFormat)
        
        // Format another range with different formatting
        let dataFormat = CellFormat.bordered(style: .thin)
        sheet.setRange("A2:C5", number: 42, format: dataFormat)
        
        // Verify formatting is applied to all cells in range
        for row in 1...1 {
            for col in 1...3 {
                let coord = CellCoordinate(row: row, column: col).excelAddress
                XCTAssertNotNil(sheet.getCellFormat(coord))
            }
        }
        
        for row in 2...5 {
            for col in 1...3 {
                let coord = CellCoordinate(row: row, column: col).excelAddress
                XCTAssertNotNil(sheet.getCellFormat(coord))
            }
        }
    }
    
    // MARK: - CSV/TSV Import/Export Tests
    
    func testCSVExport() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "CSV Test")
        
        // Add some data
        sheet.setCell("A1", string: "Name", format: CellFormat.header())
        sheet.setCell("B1", string: "Age", format: CellFormat.header())
        sheet.setCell("C1", string: "Salary", format: CellFormat.header())
        
        sheet.setCell("A2", string: "John Doe")
        sheet.setCell("B2", integer: 30)
        sheet.setCell("C2", number: 50000.50, format: CellFormat.currency())
        
        sheet.setCell("A3", string: "Jane Smith")
        sheet.setCell("B3", integer: 25)
        sheet.setCell("C3", number: 45000.75, format: CellFormat.currency())
        
        // Export to CSV
        let csv = XLKit.exportSheetToCSV(sheet: sheet)
        
        // Verify CSV content
        XCTAssertTrue(csv.contains("Name,Age,Salary"))
        XCTAssertTrue(csv.contains("John Doe,30,50000.5"))
        XCTAssertTrue(csv.contains("Jane Smith,25,45000.75"))
    }
    
    func testTSVExport() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "TSV Test")
        
        // Add some data
        sheet.setCell("A1", string: "Product")
        sheet.setCell("B1", string: "Price")
        sheet.setCell("A2", string: "Apple")
        sheet.setCell("B2", number: 1.99)
        sheet.setCell("A3", string: "Banana")
        sheet.setCell("B3", number: 0.99)
        
        // Export to TSV
        let tsv = XLKit.exportSheetToTSV(sheet: sheet)
        
        // Verify TSV content (tab-separated)
        XCTAssertTrue(tsv.contains("Product\tPrice"))
        XCTAssertTrue(tsv.contains("Apple\t1.99"))
        XCTAssertTrue(tsv.contains("Banana\t0.99"))
    }
    
    func testCSVImport() {
        let csvData = """
        Name,Age,Salary,Active
        John Doe,30,50000.50,true
        Jane Smith,25,45000.75,false
        Bob Johnson,35,60000.00,true
        """
        
        // Create workbook from CSV
        let workbook = XLKit.createWorkbookFromCSV(csvData: csvData, hasHeader: true)
        let sheet = workbook.getSheets().first!
        
        // Verify imported data (first data row is row 2)
        XCTAssertEqual(sheet.getCell("A2"), .string("John Doe"))
        XCTAssertEqual(sheet.getCell("B2"), .integer(30))
        XCTAssertEqual(sheet.getCell("C2"), .number(50000.50))
        XCTAssertEqual(sheet.getCell("D2"), .boolean(true))
        
        XCTAssertEqual(sheet.getCell("A3"), .string("Jane Smith"))
        XCTAssertEqual(sheet.getCell("B3"), .integer(25))
        XCTAssertEqual(sheet.getCell("C3"), .number(45000.75))
        XCTAssertEqual(sheet.getCell("D3"), .boolean(false))
        
        XCTAssertEqual(sheet.getCell("A4"), .string("Bob Johnson"))
        XCTAssertEqual(sheet.getCell("B4"), .integer(35))
        XCTAssertEqual(sheet.getCell("C4"), .number(60000.00))
        XCTAssertEqual(sheet.getCell("D4"), .boolean(true))
    }
    
    func testTSVImport() {
        let tsvData = """
        Product\tPrice\tIn Stock
        Apple\t1.99\ttrue
        Banana\t0.99\tfalse
        Orange\t2.49\ttrue
        """
        
        // Create workbook from TSV
        let workbook = XLKit.createWorkbookFromTSV(tsvData: tsvData, hasHeader: true)
        let sheet = workbook.getSheets().first!
        
        // Verify imported data (first data row is row 2)
        XCTAssertEqual(sheet.getCell("A2"), .string("Apple"))
        XCTAssertEqual(sheet.getCell("B2"), .number(1.99))
        XCTAssertEqual(sheet.getCell("C2"), .boolean(true))
        
        XCTAssertEqual(sheet.getCell("A3"), .string("Banana"))
        XCTAssertEqual(sheet.getCell("B3"), .number(0.99))
        XCTAssertEqual(sheet.getCell("C3"), .boolean(false))
        
        XCTAssertEqual(sheet.getCell("A4"), .string("Orange"))
        XCTAssertEqual(sheet.getCell("B4"), .number(2.49))
        XCTAssertEqual(sheet.getCell("C4"), .boolean(true))
    }
    
    func testCSVImportWithQuotes() {
        let csvData = """
        Name,Description,Value
        "John Doe","Software Engineer, Senior",50000
        "Jane Smith","Product Manager",45000
        "Bob Johnson","Data Scientist, PhD",60000
        """
        
        // Create workbook from CSV
        let workbook = XLKit.createWorkbookFromCSV(csvData: csvData, hasHeader: true)
        let sheet = workbook.getSheets().first!
        
        // Verify imported data with quotes and commas (first data row is row 2)
        XCTAssertEqual(sheet.getCell("A2"), .string("John Doe"))
        XCTAssertEqual(sheet.getCell("B2"), .string("Software Engineer, Senior"))
        XCTAssertEqual(sheet.getCell("C2"), .integer(50000))
        
        XCTAssertEqual(sheet.getCell("A3"), .string("Jane Smith"))
        XCTAssertEqual(sheet.getCell("B3"), .string("Product Manager"))
        XCTAssertEqual(sheet.getCell("C3"), .integer(45000))
        
        XCTAssertEqual(sheet.getCell("A4"), .string("Bob Johnson"))
        XCTAssertEqual(sheet.getCell("B4"), .string("Data Scientist, PhD"))
        XCTAssertEqual(sheet.getCell("C4"), .integer(60000))
    }
    
    func testCSVImportWithDates() {
        let csvData = """
        Name,Birth Date,Hire Date
        John Doe,1990-05-15,2020-03-01
        Jane Smith,1988-12-10,2019-07-15
        """
        
        // Create workbook from CSV
        let workbook = XLKit.createWorkbookFromCSV(csvData: csvData, hasHeader: true)
        let sheet = workbook.getSheets().first!
        
        // Verify imported data with dates (first data row is row 2)
        XCTAssertEqual(sheet.getCell("A2"), .string("John Doe"))
        
        // Check that dates are parsed correctly
        let birthDate1 = sheet.getCell("B2")
        if case .date(let date) = birthDate1 {
            let formatter = DateFormatter()
            formatter.dateFormat = "yyyy-MM-dd"
            XCTAssertEqual(formatter.string(from: date), "1990-05-15")
        } else {
            XCTFail("Expected date value for birth date")
        }
        
        let hireDate1 = sheet.getCell("C2")
        if case .date(let date) = hireDate1 {
            let formatter = DateFormatter()
            formatter.dateFormat = "yyyy-MM-dd"
            XCTAssertEqual(formatter.string(from: date), "2020-03-01")
        } else {
            XCTFail("Expected date value for hire date")
        }
    }
    
    func testCSVImportIntoExistingSheet() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Import Test")
        
        // Add some existing data
        sheet.setCell("A1", string: "Existing Data")
        sheet.setCell("B1", number: 100)
        
        let csvData = """
        Name,Value
        New Item 1,200
        New Item 2,300
        """
        
        // Import CSV into existing sheet
        XLKit.importCSVIntoSheet(sheet: sheet, csvData: csvData, hasHeader: true)
        
        // Verify existing data is preserved
        XCTAssertEqual(sheet.getCell("A1"), .string("Existing Data"))
        XCTAssertEqual(sheet.getCell("B1"), .number(100))
        
        // Verify new data is added (first data row is row 2)
        XCTAssertEqual(sheet.getCell("A2"), .string("New Item 1"))
        XCTAssertEqual(sheet.getCell("B2"), .integer(200))
        XCTAssertEqual(sheet.getCell("A3"), .string("New Item 2"))
        XCTAssertEqual(sheet.getCell("B3"), .integer(300))
    }
    
    func testCSVExportWithSpecialCharacters() {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Special Chars")
        
        // Add data with special characters
        sheet.setCell("A1", string: "Name")
        sheet.setCell("B1", string: "Description")
        sheet.setCell("A2", string: "John")
        sheet.setCell("B2", string: "Contains \"quotes\" and, commas")
        sheet.setCell("A3", string: "Jane")
        sheet.setCell("B3", string: "Contains\nnewlines")
        
        // Export to CSV
        let csv = XLKit.exportSheetToCSV(sheet: sheet)
        
        // Verify CSV content with proper escaping
        XCTAssertTrue(csv.contains("Name,Description"))
        XCTAssertTrue(csv.contains("John,\"Contains \"\"quotes\"\" and, commas\""))
        XCTAssertTrue(csv.contains("Jane,\"Contains\nnewlines\""))
    }
    
    func testExcelCellDimensionsForAnyImageSize() async throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Cell Dimensions Test")
        
        // Test with various image dimensions including movie/video aspect ratios
        let testCases = [
            (createWidePNG(), "A1", "16:9 wide"),
            (createSquarePNG(), "B1", "1:1 square"),
            (createTallPNG(), "C1", "9:16 tall"),
            (createUltraWidePNG(), "D1", "21:9 ultra-wide"),
            (createPortraitPNG(), "E1", "3:4 portrait"),
            (createCinemascopePNG(), "F1", "2.39:1 cinemascope"),
            (createAcademyPNG(), "G1", "1.85:1 academy"),
            (createClassicTVPNG(), "H1", "4:3 classic TV"),
            (createModernMobilePNG(), "I1", "18:9 modern mobile")
        ]
        
        for (imageData, coordinate, description) in testCases {
            // Embed image with auto-sizing
            let success = sheet.embedImage(
                imageData,
                at: coordinate,
                of: workbook,
                scale: 0.5, // 50% of max size
                maxWidth: 400,
                maxHeight: 300
            )
            
            XCTAssertTrue(success, "Failed to embed \(description)")
            
            // Get the embedded image
            let images = sheet.getImages()
            guard let image = images[coordinate] else {
                XCTFail("Image not found for \(description)")
                continue
            }
            
            // Get cell coordinate
            guard let cellCoord = CellCoordinate(excelAddress: coordinate) else {
                XCTFail("Invalid coordinate for \(description)")
                continue
            }
            
            // Get Excel cell dimensions
            let colWidth = sheet.getColumnWidth(cellCoord.column)
            let rowHeight = sheet.getRowHeight(cellCoord.row)
            
            XCTAssertNotNil(colWidth, "Column width not set for \(description)")
            XCTAssertNotNil(rowHeight, "Row height not set for \(description)")
            
            // Calculate expected Excel dimensions from display size using new formulas
            let expectedColWidth = ImageSizingUtils.excelColumnWidth(forPixelWidth: image.displaySize!.width)
            let expectedRowHeight = ImageSizingUtils.excelRowHeight(forPixelHeight: image.displaySize!.height)
            
            // Verify Excel dimensions match expected values
            XCTAssertEqual(colWidth!, expectedColWidth, accuracy: 0.001, 
                          "Column width mismatch for \(description)")
            XCTAssertEqual(rowHeight!, expectedRowHeight, accuracy: 0.001, 
                          "Row height mismatch for \(description)")
            
            // Verify aspect ratio is preserved in Excel cell dimensions
            let imageAspectRatio = image.displaySize!.width / image.displaySize!.height
            let cellAspectRatio = (colWidth! * 8.0) / (rowHeight! * 1.33)
            XCTAssertEqual(imageAspectRatio, cellAspectRatio, accuracy: 0.1, 
                          "Cell aspect ratio doesn't match image for \(description)")
            
            print("[TEST] \(description): image=\(image.displaySize!), col=\(colWidth!), row=\(rowHeight!), aspect=\(imageAspectRatio)")
        }
        
        // Save and validate the workbook
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("CellDimensionsTest.xlsx")
        try await XLKit.saveWorkbook(workbook, to: tempURL)
        
        // Verify file was created
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
        
        // Clean up
        try FileManager.default.removeItem(at: tempURL)
    }
} 