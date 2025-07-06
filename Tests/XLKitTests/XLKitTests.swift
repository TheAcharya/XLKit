import XCTest
@testable import XLKit

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
yes        let sheet1 = workbook.addSheet(name: "Sheet1")
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
        
        sheet.setRange("A1:C3", value: .string("Range"))
        
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
        XCTAssertEqual(XLKitUtils.columnLetter(from: 1), "A")
        XCTAssertEqual(XLKitUtils.columnLetter(from: 2), "B")
        XCTAssertEqual(XLKitUtils.columnLetter(from: 26), "Z")
        XCTAssertEqual(XLKitUtils.columnLetter(from: 27), "AA")
        XCTAssertEqual(XLKitUtils.columnLetter(from: 28), "AB")
        
        XCTAssertEqual(XLKitUtils.columnNumber(from: "A"), 1)
        XCTAssertEqual(XLKitUtils.columnNumber(from: "B"), 2)
        XCTAssertEqual(XLKitUtils.columnNumber(from: "Z"), 26)
        XCTAssertEqual(XLKitUtils.columnNumber(from: "AA"), 27)
        XCTAssertEqual(XLKitUtils.columnNumber(from: "AB"), 28)
    }
    
    func testDateConversion() {
        let date = Date()
        let excelNumber = XLKitUtils.excelNumberFromDate(date)
        let convertedDate = XLKitUtils.dateFromExcelNumber(excelNumber)
        
        // Allow for small precision differences
        XCTAssertEqual(date.timeIntervalSince1970, convertedDate.timeIntervalSince1970, accuracy: 1.0)
    }
    
    func testXMLEscaping() {
        XCTAssertEqual(XLKitUtils.escapeXML("&"), "&amp;")
        XCTAssertEqual(XLKitUtils.escapeXML("<"), "&lt;")
        XCTAssertEqual(XLKitUtils.escapeXML(">"), "&gt;")
        XCTAssertEqual(XLKitUtils.escapeXML("\""), "&quot;")
        XCTAssertEqual(XLKitUtils.escapeXML("'"), "&apos;")
        XCTAssertEqual(XLKitUtils.escapeXML("Hello & World"), "Hello &amp; World")
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
        XCTAssertEqual(XLKitUtils.detectImageFormat(from: gifData), .gif)
        
        // Test PNG format detection
        let pngData = Data([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])
        XCTAssertEqual(XLKitUtils.detectImageFormat(from: pngData), .png)
        
        // Test JPEG format detection
        let jpegData = Data([0xFF, 0xD8, 0xFF, 0xE0])
        XCTAssertEqual(XLKitUtils.detectImageFormat(from: jpegData), .jpeg)
        
        // Test invalid data
        let invalidData = Data([0x00, 0x00, 0x00, 0x00])
        XCTAssertNil(XLKitUtils.detectImageFormat(from: invalidData))
    }
    
    func testImageSizeDetection() {
        // Test GIF size detection
        let gifData = Data([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00])
        let gifSize = XLKitUtils.getImageSize(from: gifData, format: .gif)
        XCTAssertEqual(gifSize?.width, 16)
        XCTAssertEqual(gifSize?.height, 32)
        
        // Test PNG size detection
        let pngData = Data([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x00, 0x30])
        let pngSize = XLKitUtils.getImageSize(from: pngData, format: .png)
        XCTAssertEqual(pngSize?.width, 64)
        XCTAssertEqual(pngSize?.height, 48)
    }
    
    func testExcelImageCreation() {
        let imageData = Data([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00])
        
        // Test creation from data
        let image = ExcelImage.from(data: imageData, format: .gif)
        XCTAssertNotNil(image)
        XCTAssertEqual(image?.format, .gif)
        XCTAssertEqual(image?.originalSize.width, 16)
        XCTAssertEqual(image?.originalSize.height, 32)
        
        // Test creation with display size
        let imageWithSize = ExcelImage.from(data: imageData, format: .gif, displaySize: CGSize(width: 100, height: 200))
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
        let image = ExcelImage.from(data: imageData, format: .gif)!
        
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
} 