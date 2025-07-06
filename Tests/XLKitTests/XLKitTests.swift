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
        _ = workbook.addSheet(name: "Sheet1")
        _ = workbook.addSheet(name: "Sheet2")
        
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
} 