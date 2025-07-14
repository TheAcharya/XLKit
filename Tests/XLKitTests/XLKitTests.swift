//
//  XLKitTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class XLKitTests: XCTestCase {
    
    func testCreateWorkbook() throws {
        let workbook = Workbook()
        XCTAssertNotNil(workbook)
        XCTAssertEqual(workbook.getSheets().count, 0)
    }
    
    func testAddSheet() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test Sheet")
        
        XCTAssertEqual(workbook.getSheets().count, 1)
        XCTAssertEqual(sheet.name, "Test Sheet")
        XCTAssertEqual(sheet.id, 1)
    }
    
    func testGetSheetByName() throws {
        let workbook = Workbook()
        let sheet1 = workbook.addSheet(name: "Sheet1")
        let sheet2 = workbook.addSheet(name: "Sheet2")
        
        XCTAssertEqual(workbook.getSheet(name: "Sheet1"), sheet1)
        XCTAssertEqual(workbook.getSheet(name: "Sheet2"), sheet2)
        XCTAssertNil(workbook.getSheet(name: "NonExistent"))
    }
    
    func testRemoveSheet() throws {
        let workbook = Workbook()
        _ = workbook.addSheet(name: "Sheet1")
        _ = workbook.addSheet(name: "Sheet2")
        
        XCTAssertEqual(workbook.getSheets().count, 2)
        
        workbook.removeSheet(name: "Sheet1")
        XCTAssertEqual(workbook.getSheets().count, 1)
        XCTAssertNil(workbook.getSheet(name: "Sheet1"))
        XCTAssertNotNil(workbook.getSheet(name: "Sheet2"))
    }
    
    func testSetAndGetCell() throws {
        let workbook = Workbook()
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
    
    func testSetCellByRowColumn() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell(row: 1, column: 1, value: .string("A1"))
        sheet.setCell(row: 2, column: 2, value: .string("B2"))
        
        XCTAssertEqual(sheet.getCell("A1"), .string("A1"))
        XCTAssertEqual(sheet.getCell("B2"), .string("B2"))
    }
    
    func testSetRange() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRange("A1:C3", string: "Range")
        
        XCTAssertEqual(sheet.getCell("A1"), .string("Range"))
        XCTAssertEqual(sheet.getCell("B2"), .string("Range"))
        XCTAssertEqual(sheet.getCell("C3"), .string("Range"))
    }
    
    func testMergeCells() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.mergeCells("A1:C1")
        
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, 1)
        XCTAssertEqual(mergedRanges[0].excelRange, "A1:C1")
    }
    
    func testGetUsedCells() throws {
        let workbook = Workbook()
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
    
    func testClearSheet() throws {
        let workbook = Workbook()
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
        XCTAssertEqual(CellValue.formula("=A1").type, "formula")
        XCTAssertEqual(CellValue.empty.type, "empty")
    }
    
    func testCellFormatting() {
        let headerFormat = CellFormat.header()
        let currencyFormat = CellFormat.currency()
        let percentageFormat = CellFormat.percentage()
        let dateFormat = CellFormat.date()
        let borderedFormat = CellFormat.bordered()
        
        XCTAssertNotNil(headerFormat)
        XCTAssertNotNil(currencyFormat)
        XCTAssertNotNil(percentageFormat)
        XCTAssertNotNil(dateFormat)
        XCTAssertNotNil(borderedFormat)
    }
    
    func testSetCellWithFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let headerCell = Cell.string("Header", format: CellFormat.header())
        let currencyCell = Cell.number(1234.56, format: CellFormat.currency())
        
        sheet.setCell("A1", cell: headerCell)
        sheet.setCell("B1", cell: currencyCell)
        
        let retrievedHeader = sheet.getCellWithFormat("A1")
        let retrievedCurrency = sheet.getCellWithFormat("B1")
        
        XCTAssertNotNil(retrievedHeader)
        XCTAssertNotNil(retrievedCurrency)
        XCTAssertEqual(retrievedHeader?.value, .string("Header"))
        XCTAssertEqual(retrievedCurrency?.value, .number(1234.56))
    }
    
    func testColumnAndRowSizing() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setColumnWidth(1, width: 100.0)
        sheet.setRowHeight(1, height: 25.0)
        
        XCTAssertEqual(sheet.getColumnWidth(1), 100.0)
        XCTAssertEqual(sheet.getRowHeight(1), 25.0)
    }
    
    func testWorkbookImageManagement() {
        let workbook = Workbook()
        
        let imageData = Data()
        let image = ExcelImage(
            id: "test-image",
            data: imageData,
            format: .png,
            originalSize: CGSize(width: 100, height: 100)
        )
        
        workbook.addImage(image)
        XCTAssertEqual(workbook.imageCount, 1)
        
        let retrievedImage = workbook.getImage(withId: "test-image")
        XCTAssertNotNil(retrievedImage)
        XCTAssertEqual(retrievedImage?.id, "test-image")
        
        workbook.removeImage(withId: "test-image")
        XCTAssertEqual(workbook.imageCount, 0)
    }
    
    func testSheetImageManagement() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let imageData = Data()
        let image = ExcelImage(
            id: "test-image",
            data: imageData,
            format: .png,
            originalSize: CGSize(width: 100, height: 100)
        )
        
        sheet.addImage(image, at: "A1")
        XCTAssertTrue(sheet.hasImage(at: "A1"))
        
        let retrievedImage = sheet.getImage(at: "A1")
        XCTAssertNotNil(retrievedImage)
        XCTAssertEqual(retrievedImage?.id, "test-image")
        
        sheet.removeImage(at: "A1")
        XCTAssertFalse(sheet.hasImage(at: "A1"))
    }
    
    func testWorkbookSave() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        sheet.setCell("A1", value: .string("Test"))
        
        let tempURL = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("test.xlsx")
        
        // Test synchronous save
        try workbook.save(to: tempURL)
        
        // Verify file exists
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
        
        // Clean up
        try FileManager.default.removeItem(at: tempURL)
    }
    
    func testWorkbookSaveAsync() async throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        sheet.setCell("A1", value: .string("Test"))
        
        let tempURL = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("test_async.xlsx")
        
        // Test asynchronous save
        try await workbook.save(to: tempURL)
        
        // Verify file exists
        XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
        
        // Clean up
        try FileManager.default.removeItem(at: tempURL)
    }
    
    func testWorkbookFromCSV() {
        let csvData = """
        Name,Age,Salary
        Alice,30,50000.5
        Bob,25,45000.75
        """
        
        let workbook = Workbook(fromCSV: csvData, hasHeader: true)
        let sheet = workbook.getSheets().first!
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Alice"))
        XCTAssertEqual(sheet.getCell("B2"), .integer(30))
        XCTAssertEqual(sheet.getCell("C2"), .number(50000.5))
        XCTAssertEqual(sheet.getCell("A3"), .string("Bob"))
        XCTAssertEqual(sheet.getCell("B3"), .integer(25))
        XCTAssertEqual(sheet.getCell("C3"), .number(45000.75))
    }
    
    func testWorkbookFromTSV() {
        let tsvData = """
        Product\tPrice\tIn Stock
        Apple\t1.99\ttrue
        Banana\t0.99\tfalse
        """
        
        let workbook = Workbook(fromTSV: tsvData, hasHeader: true)
        let sheet = workbook.getSheets().first!
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Apple"))
        XCTAssertEqual(sheet.getCell("B2"), .number(1.99))
        XCTAssertEqual(sheet.getCell("C2"), .boolean(true))
        XCTAssertEqual(sheet.getCell("A3"), .string("Banana"))
        XCTAssertEqual(sheet.getCell("B3"), .number(0.99))
        XCTAssertEqual(sheet.getCell("C3"), .boolean(false))
    }
    
    func testSheetExportToCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Name", "Age", "Salary"])
        sheet.setRow(2, strings: ["Alice", "30", "50000"])
        sheet.setRow(3, strings: ["Bob", "25", "45000"])
        
        let csv = sheet.exportToCSV()
        let expectedCSV = "Name,Age,Salary\nAlice,30,50000\nBob,25,45000"
        
        XCTAssertEqual(csv, expectedCSV)
    }
    
    func testSheetExportToTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Product", "Price", "In Stock"])
        sheet.setRow(2, strings: ["Apple", "1.99", "true"])
        sheet.setRow(3, strings: ["Banana", "0.99", "false"])
        
        let tsv = sheet.exportToTSV()
        let expectedTSV = "Product\tPrice\tIn Stock\nApple\t1.99\ttrue\nBanana\t0.99\tfalse"
        
        XCTAssertEqual(tsv, expectedTSV)
    }
    
    func testWorkbookImportCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let csvData = """
        Name,Age,Salary
        Alice,30,50000.5
        Bob,25,45000.75
        """
        
        workbook.importCSV(csvData, into: sheet, hasHeader: true)
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Alice"))
        XCTAssertEqual(sheet.getCell("B2"), .integer(30))
        XCTAssertEqual(sheet.getCell("C2"), .number(50000.5))
        XCTAssertEqual(sheet.getCell("A3"), .string("Bob"))
        XCTAssertEqual(sheet.getCell("B3"), .integer(25))
        XCTAssertEqual(sheet.getCell("C3"), .number(45000.75))
    }
    
    func testWorkbookImportTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        let tsvData = """
        Product\tPrice\tIn Stock
        Apple\t1.99\ttrue
        Banana\t0.99\tfalse
        """
        
        workbook.importTSV(tsvData, into: sheet, hasHeader: true)
        
        XCTAssertEqual(sheet.getCell("A2"), .string("Apple"))
        XCTAssertEqual(sheet.getCell("B2"), .number(1.99))
        XCTAssertEqual(sheet.getCell("C2"), .boolean(true))
        XCTAssertEqual(sheet.getCell("A3"), .string("Banana"))
        XCTAssertEqual(sheet.getCell("B3"), .number(0.99))
        XCTAssertEqual(sheet.getCell("C3"), .boolean(false))
    }
    
    func testWorkbookExportSheetToCSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Name", "Age", "Salary"])
        sheet.setRow(2, strings: ["Alice", "30", "50000"])
        sheet.setRow(3, strings: ["Bob", "25", "45000"])
        
        let csv = workbook.exportSheetToCSV(sheet)
        let expectedCSV = "Name,Age,Salary\nAlice,30,50000\nBob,25,45000"
        
        XCTAssertEqual(csv, expectedCSV)
    }
    
    func testWorkbookExportSheetToTSV() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRow(1, strings: ["Product", "Price", "In Stock"])
        sheet.setRow(2, strings: ["Apple", "1.99", "true"])
        sheet.setRow(3, strings: ["Banana", "0.99", "false"])
        
        let tsv = workbook.exportSheetToTSV(sheet)
        let expectedTSV = "Product\tPrice\tIn Stock\nApple\t1.99\ttrue\nBanana\t0.99\tfalse"
        
        XCTAssertEqual(tsv, expectedTSV)
    }
    
    func testFluentAPI() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test fluent API with method chaining
        sheet
            .setCell("A1", value: .string("Header"))
            .setCell("B1", value: .string("Value"))
            .setRow(2, strings: ["Row1", "Data1"])
            .setRow(3, strings: ["Row2", "Data2"])
            .mergeCells("A1:B1")
        
        XCTAssertEqual(sheet.getCell("A1"), .string("Header"))
        XCTAssertEqual(sheet.getCell("B1"), .string("Value"))
        XCTAssertEqual(sheet.getCell("A2"), .string("Row1"))
        XCTAssertEqual(sheet.getCell("B2"), .string("Data1"))
        XCTAssertEqual(sheet.getCell("A3"), .string("Row2"))
        XCTAssertEqual(sheet.getCell("B3"), .string("Data2"))
        
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, 1)
        XCTAssertEqual(mergedRanges[0].excelRange, "A1:B1")
    }
    
    func testSheetConvenienceMethods() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test convenience methods
        sheet.setCell("A1", string: "String Value")
        sheet.setCell("B1", number: 123.45)
        sheet.setCell("C1", integer: 42)
        sheet.setCell("D1", boolean: true)
        sheet.setCell("E1", date: Date())
        sheet.setCell("F1", formula: "=A1+B1")
        
        XCTAssertEqual(sheet.getCell("A1"), .string("String Value"))
        XCTAssertEqual(sheet.getCell("B1"), .number(123.45))
        XCTAssertEqual(sheet.getCell("C1"), .integer(42))
        XCTAssertEqual(sheet.getCell("D1"), .boolean(true))
        XCTAssertEqual(sheet.getCell("E1")?.type, "date")
        XCTAssertEqual(sheet.getCell("F1"), .formula("=A1+B1"))
    }
    
    func testSheetRowAndColumnMethods() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test row methods
        sheet.setRow(1, strings: ["Name", "Age", "Salary"])
        sheet.setRow(2, numbers: [30.0, 25.0, 50000.0])
        sheet.setRow(3, integers: [1, 2, 3])
        
        // Test column methods
        sheet.setColumn(4, strings: ["Col1", "Col2", "Col3"])
        sheet.setColumn(5, numbers: [10.5, 20.5, 30.5])
        sheet.setColumn(6, integers: [100, 200, 300])
        
        XCTAssertEqual(sheet.getCell("A1"), .string("Name"))
        XCTAssertEqual(sheet.getCell("B1"), .string("Age"))
        XCTAssertEqual(sheet.getCell("C1"), .string("Salary"))
        XCTAssertEqual(sheet.getCell("A2"), .number(30.0))
        XCTAssertEqual(sheet.getCell("B2"), .number(25.0))
        XCTAssertEqual(sheet.getCell("C2"), .number(50000.0))
        XCTAssertEqual(sheet.getCell("A3"), .integer(1))
        XCTAssertEqual(sheet.getCell("B3"), .integer(2))
        XCTAssertEqual(sheet.getCell("C3"), .integer(3))
        
        XCTAssertEqual(sheet.getCell("D1"), .string("Col1"))
        XCTAssertEqual(sheet.getCell("D2"), .string("Col2"))
        XCTAssertEqual(sheet.getCell("D3"), .string("Col3"))
        XCTAssertEqual(sheet.getCell("E1"), .number(10.5))
        XCTAssertEqual(sheet.getCell("E2"), .number(20.5))
        XCTAssertEqual(sheet.getCell("E3"), .number(30.5))
        XCTAssertEqual(sheet.getCell("F1"), .integer(100))
        XCTAssertEqual(sheet.getCell("F2"), .integer(200))
        XCTAssertEqual(sheet.getCell("F3"), .integer(300))
    }
    
    func testSheetUtilityProperties() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test empty sheet
        XCTAssertTrue(sheet.isEmpty)
        XCTAssertEqual(sheet.cellCount, 0)
        XCTAssertEqual(sheet.imageCount, 0)
        
        // Add some data
        sheet.setCell("A1", value: .string("Test"))
        sheet.setCell("B2", value: .number(42.5))
        
        XCTAssertFalse(sheet.isEmpty)
        XCTAssertEqual(sheet.cellCount, 2)
        
        // Test allCells property
        let allCells = sheet.allCells
        XCTAssertEqual(allCells.count, 2)
        XCTAssertEqual(allCells["A1"], .string("Test"))
        XCTAssertEqual(allCells["B2"], .number(42.5))
        
        // Test allFormattedCells property
        let allFormattedCells = sheet.allFormattedCells
        XCTAssertEqual(allFormattedCells.count, 2)
        XCTAssertEqual(allFormattedCells["A1"]?.value, .string("Test"))
        XCTAssertEqual(allFormattedCells["B2"]?.value, .number(42.5))
    }
    
    func testSheetConvenienceInitializer() {
        let initialData = [
            "A1": CellValue.string("Header"),
            "B1": CellValue.string("Value"),
            "A2": CellValue.number(42.5)
        ]
        
        let sheet = Sheet(name: "Test", id: 1)
        
        // Set the initial data manually
        for (coordinate, value) in initialData {
            sheet.setCell(coordinate, value: value)
        }
        
        XCTAssertEqual(sheet.getCell("A1"), CellValue.string("Header"))
        XCTAssertEqual(sheet.getCell("B1"), CellValue.string("Value"))
        XCTAssertEqual(sheet.getCell("A2"), CellValue.number(42.5))
        XCTAssertEqual(sheet.cellCount, 3)
    }
} 