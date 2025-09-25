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
    
    func testFontColorFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Font Color Test")
        
        // Test different font colors
        let redTextFormat = CellFormat.text(color: "#FF0000")
        let blueTextFormat = CellFormat.text(color: "#0000FF")
        let greenTextFormat = CellFormat.text(color: "#00FF00")
        let currencyRedFormat = CellFormat.currency(color: "#FF0000")
        
        // Set cells with different colors
        sheet.setCell("A1", string: "Red Text", format: redTextFormat)
        sheet.setCell("A2", string: "Blue Text", format: blueTextFormat)
        sheet.setCell("A3", string: "Green Text", format: greenTextFormat)
        sheet.setCell("A4", number: 1234.56, format: currencyRedFormat)
        
        // Verify formats are applied
        let redCell = sheet.getCellWithFormat("A1")
        let blueCell = sheet.getCellWithFormat("A2")
        let greenCell = sheet.getCellWithFormat("A3")
        let currencyCell = sheet.getCellWithFormat("A4")
        
        XCTAssertNotNil(redCell)
        XCTAssertNotNil(blueCell)
        XCTAssertNotNil(greenCell)
        XCTAssertNotNil(currencyCell)
        
        XCTAssertEqual(redCell?.format?.fontColor, "#FF0000")
        XCTAssertEqual(blueCell?.format?.fontColor, "#0000FF")
        XCTAssertEqual(greenCell?.format?.fontColor, "#00FF00")
        XCTAssertEqual(currencyCell?.format?.fontColor, "#FF0000")
        
        // Test that the format key includes font color
        if let redFormat = redCell?.format {
            let formatKey = XLSXEngine.formatToKey(redFormat)
            XCTAssertTrue(formatKey.contains("fontColor:#FF0000"))
        } else {
            XCTFail("Red cell format should not be nil")
        }
    }
    
    func testHorizontalAlignment() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Horizontal Alignment Test")
        
        // Test all horizontal alignment options
        var leftAlignFormat = CellFormat()
        leftAlignFormat.horizontalAlignment = .left
        
        var centerAlignFormat = CellFormat()
        centerAlignFormat.horizontalAlignment = .center
        
        var rightAlignFormat = CellFormat()
        rightAlignFormat.horizontalAlignment = .right
        
        var justifyAlignFormat = CellFormat()
        justifyAlignFormat.horizontalAlignment = .justify
        
        var distributedAlignFormat = CellFormat()
        distributedAlignFormat.horizontalAlignment = .distributed
        
        // Set cells with different horizontal alignments
        sheet.setCell("A1", string: "Left Aligned", format: leftAlignFormat)
        sheet.setCell("A2", string: "Center Aligned", format: centerAlignFormat)
        sheet.setCell("A3", string: "Right Aligned", format: rightAlignFormat)
        sheet.setCell("A4", string: "Justify Aligned", format: justifyAlignFormat)
        sheet.setCell("A5", string: "Distributed Aligned", format: distributedAlignFormat)
        
        // Verify alignments are applied
        let leftCell = sheet.getCellWithFormat("A1")
        let centerCell = sheet.getCellWithFormat("A2")
        let rightCell = sheet.getCellWithFormat("A3")
        let justifyCell = sheet.getCellWithFormat("A4")
        let distributedCell = sheet.getCellWithFormat("A5")
        
        XCTAssertNotNil(leftCell)
        XCTAssertNotNil(centerCell)
        XCTAssertNotNil(rightCell)
        XCTAssertNotNil(justifyCell)
        XCTAssertNotNil(distributedCell)
        
        XCTAssertEqual(leftCell?.format?.horizontalAlignment, .left)
        XCTAssertEqual(centerCell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(rightCell?.format?.horizontalAlignment, .right)
        XCTAssertEqual(justifyCell?.format?.horizontalAlignment, .justify)
        XCTAssertEqual(distributedCell?.format?.horizontalAlignment, .distributed)
        
        // Test that format keys include horizontal alignment
        if let leftFormat = leftCell?.format {
            let formatKey = XLSXEngine.formatToKey(leftFormat)
            XCTAssertTrue(formatKey.contains("horizontalAlignment:left"))
        }
        
        if let centerFormat = centerCell?.format {
            let formatKey = XLSXEngine.formatToKey(centerFormat)
            XCTAssertTrue(formatKey.contains("horizontalAlignment:center"))
        }
    }
    
    func testVerticalAlignment() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Vertical Alignment Test")
        
        // Test all vertical alignment options
        var topAlignFormat = CellFormat()
        topAlignFormat.verticalAlignment = .top
        
        var centerAlignFormat = CellFormat()
        centerAlignFormat.verticalAlignment = .center
        
        var bottomAlignFormat = CellFormat()
        bottomAlignFormat.verticalAlignment = .bottom
        
        var justifyAlignFormat = CellFormat()
        justifyAlignFormat.verticalAlignment = .justify
        
        var distributedAlignFormat = CellFormat()
        distributedAlignFormat.verticalAlignment = .distributed
        
        // Set cells with different vertical alignments
        sheet.setCell("A1", string: "Top Aligned", format: topAlignFormat)
        sheet.setCell("A2", string: "Center Aligned", format: centerAlignFormat)
        sheet.setCell("A3", string: "Bottom Aligned", format: bottomAlignFormat)
        sheet.setCell("A4", string: "Justify Aligned", format: justifyAlignFormat)
        sheet.setCell("A5", string: "Distributed Aligned", format: distributedAlignFormat)
        
        // Verify alignments are applied
        let topCell = sheet.getCellWithFormat("A1")
        let centerCell = sheet.getCellWithFormat("A2")
        let bottomCell = sheet.getCellWithFormat("A3")
        let justifyCell = sheet.getCellWithFormat("A4")
        let distributedCell = sheet.getCellWithFormat("A5")
        
        XCTAssertNotNil(topCell)
        XCTAssertNotNil(centerCell)
        XCTAssertNotNil(bottomCell)
        XCTAssertNotNil(justifyCell)
        XCTAssertNotNil(distributedCell)
        
        XCTAssertEqual(topCell?.format?.verticalAlignment, .top)
        XCTAssertEqual(centerCell?.format?.verticalAlignment, .center)
        XCTAssertEqual(bottomCell?.format?.verticalAlignment, .bottom)
        XCTAssertEqual(justifyCell?.format?.verticalAlignment, .justify)
        XCTAssertEqual(distributedCell?.format?.verticalAlignment, .distributed)
        
        // Test that format keys include vertical alignment
        if let topFormat = topCell?.format {
            let formatKey = XLSXEngine.formatToKey(topFormat)
            XCTAssertTrue(formatKey.contains("verticalAlignment:top"))
        }
        
        if let centerFormat = centerCell?.format {
            let formatKey = XLSXEngine.formatToKey(centerFormat)
            XCTAssertTrue(formatKey.contains("verticalAlignment:center"))
        }
    }
    
    func testCombinedAlignment() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Combined Alignment Test")
        
        // Test combinations of horizontal and vertical alignments
        var topLeftFormat = CellFormat()
        topLeftFormat.horizontalAlignment = .left
        topLeftFormat.verticalAlignment = .top
        
        var centerCenterFormat = CellFormat()
        centerCenterFormat.horizontalAlignment = .center
        centerCenterFormat.verticalAlignment = .center
        
        var bottomRightFormat = CellFormat()
        bottomRightFormat.horizontalAlignment = .right
        bottomRightFormat.verticalAlignment = .bottom
        
        var justifyDistributedFormat = CellFormat()
        justifyDistributedFormat.horizontalAlignment = .justify
        justifyDistributedFormat.verticalAlignment = .distributed
        
        // Set cells with combined alignments
        sheet.setCell("A1", string: "Top-Left", format: topLeftFormat)
        sheet.setCell("A2", string: "Center-Center", format: centerCenterFormat)
        sheet.setCell("A3", string: "Bottom-Right", format: bottomRightFormat)
        sheet.setCell("A4", string: "Justify-Distributed", format: justifyDistributedFormat)
        
        // Verify combined alignments are applied
        let topLeftCell = sheet.getCellWithFormat("A1")
        let centerCenterCell = sheet.getCellWithFormat("A2")
        let bottomRightCell = sheet.getCellWithFormat("A3")
        let justifyDistributedCell = sheet.getCellWithFormat("A4")
        
        XCTAssertNotNil(topLeftCell)
        XCTAssertNotNil(centerCenterCell)
        XCTAssertNotNil(bottomRightCell)
        XCTAssertNotNil(justifyDistributedCell)
        
        // Test horizontal alignments
        XCTAssertEqual(topLeftCell?.format?.horizontalAlignment, .left)
        XCTAssertEqual(centerCenterCell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(bottomRightCell?.format?.horizontalAlignment, .right)
        XCTAssertEqual(justifyDistributedCell?.format?.horizontalAlignment, .justify)
        
        // Test vertical alignments
        XCTAssertEqual(topLeftCell?.format?.verticalAlignment, .top)
        XCTAssertEqual(centerCenterCell?.format?.verticalAlignment, .center)
        XCTAssertEqual(bottomRightCell?.format?.verticalAlignment, .bottom)
        XCTAssertEqual(justifyDistributedCell?.format?.verticalAlignment, .distributed)
        
        // Test that format keys include both alignments
        if let topLeftFormat = topLeftCell?.format {
            let formatKey = XLSXEngine.formatToKey(topLeftFormat)
            XCTAssertTrue(formatKey.contains("horizontalAlignment:left"))
            XCTAssertTrue(formatKey.contains("verticalAlignment:top"))
        }
        
        if let centerCenterFormat = centerCenterCell?.format {
            let formatKey = XLSXEngine.formatToKey(centerCenterFormat)
            XCTAssertTrue(formatKey.contains("horizontalAlignment:center"))
            XCTAssertTrue(formatKey.contains("verticalAlignment:center"))
        }
    }
    
    func testAlignmentWithOtherFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Alignment with Formatting Test")
        
        // Test alignment combined with other formatting options
        var formattedAlignFormat = CellFormat()
        formattedAlignFormat.horizontalAlignment = .center
        formattedAlignFormat.verticalAlignment = .center
        formattedAlignFormat.fontWeight = .bold
        formattedAlignFormat.fontSize = 14.0
        formattedAlignFormat.fontColor = "#FF0000"
        formattedAlignFormat.backgroundColor = "#FFFF00"
        
        // Set cell with combined formatting
        sheet.setCell("A1", string: "Formatted & Aligned", format: formattedAlignFormat)
        
        // Verify all formatting is applied
        let formattedCell = sheet.getCellWithFormat("A1")
        XCTAssertNotNil(formattedCell)
        
        XCTAssertEqual(formattedCell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(formattedCell?.format?.verticalAlignment, .center)
        XCTAssertEqual(formattedCell?.format?.fontWeight, .bold)
        XCTAssertEqual(formattedCell?.format?.fontSize, 14.0)
        XCTAssertEqual(formattedCell?.format?.fontColor, "#FF0000")
        XCTAssertEqual(formattedCell?.format?.backgroundColor, "#FFFF00")
        
        // Test that format key includes all formatting
        if let format = formattedCell?.format {
            let formatKey = XLSXEngine.formatToKey(format)
            XCTAssertTrue(formatKey.contains("horizontalAlignment:center"))
            XCTAssertTrue(formatKey.contains("verticalAlignment:center"))
            XCTAssertTrue(formatKey.contains("fontWeight:bold"))
            XCTAssertTrue(formatKey.contains("fontSize:14.0"))
            XCTAssertTrue(formatKey.contains("fontColor:#FF0000"))
            XCTAssertTrue(formatKey.contains("backgroundColor:#FFFF00"))
        }
    }
    
    func testAlignmentEnumValues() {
        // Test that all alignment enum values are correct
        XCTAssertEqual(HorizontalAlignment.left.rawValue, "left")
        XCTAssertEqual(HorizontalAlignment.center.rawValue, "center")
        XCTAssertEqual(HorizontalAlignment.right.rawValue, "right")
        XCTAssertEqual(HorizontalAlignment.justify.rawValue, "justify")
        XCTAssertEqual(HorizontalAlignment.distributed.rawValue, "distributed")
        
        XCTAssertEqual(VerticalAlignment.top.rawValue, "top")
        XCTAssertEqual(VerticalAlignment.center.rawValue, "center")
        XCTAssertEqual(VerticalAlignment.bottom.rawValue, "bottom")
        XCTAssertEqual(VerticalAlignment.justify.rawValue, "justify")
        XCTAssertEqual(VerticalAlignment.distributed.rawValue, "distributed")
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
    
    // MARK: - Number Format Tests
    
    func testNumberFormatCurrency() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Test currency formatting
        let currencyFormat = CellFormat.currency()
        sheet.setCell("A1", number: 1234.56, format: currencyFormat)
        
        // Verify the format is stored correctly
        let storedFormat = sheet.getCellFormat("A1")
        XCTAssertNotNil(storedFormat)
        XCTAssertEqual(storedFormat?.numberFormat, .currencyWithDecimals)
        
        // Test custom currency format
        var customCurrencyFormat = CellFormat.currency()
        customCurrencyFormat.numberFormat = .custom
        customCurrencyFormat.customNumberFormat = "$#,##0.00"
        sheet.setCell("A2", number: 5678.90, format: customCurrencyFormat)
        
        let storedCustomFormat = sheet.getCellFormat("A2")
        XCTAssertNotNil(storedCustomFormat)
        XCTAssertEqual(storedCustomFormat?.numberFormat, .custom)
        XCTAssertEqual(storedCustomFormat?.customNumberFormat, "$#,##0.00")
    }
    
    func testNumberFormatPercentage() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Test percentage formatting
        let percentageFormat = CellFormat.percentage()
        sheet.setCell("A1", number: 0.1234, format: percentageFormat)
        
        // Verify the format is stored correctly
        let storedFormat = sheet.getCellFormat("A1")
        XCTAssertNotNil(storedFormat)
        XCTAssertEqual(storedFormat?.numberFormat, .percentageWithDecimals)
    }
    
    func testNumberFormatCustom() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Test custom number format
        var customFormat = CellFormat()
        customFormat.numberFormat = .custom
        customFormat.customNumberFormat = "#,##0.00"
        
        sheet.setCell("A1", number: 1234.56, format: customFormat)
        
        // Verify the format is stored correctly
        let storedFormat = sheet.getCellFormat("A1")
        XCTAssertNotNil(storedFormat)
        XCTAssertEqual(storedFormat?.numberFormat, .custom)
        XCTAssertEqual(storedFormat?.customNumberFormat, "#,##0.00")
    }
    
    func testNumberFormatInFormatToKey() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Create formats with different number formats
        let currencyFormat = CellFormat.currency()
        let percentageFormat = CellFormat.percentage()
        var customFormat = CellFormat()
        customFormat.numberFormat = .custom
        customFormat.customNumberFormat = "#,##0.00"
        
        // Set cells with different formats
        sheet.setCell("A1", number: 1234.56, format: currencyFormat)
        sheet.setCell("A2", number: 0.1234, format: percentageFormat)
        sheet.setCell("A3", number: 5678.90, format: customFormat)
        
        // Verify that formatToKey generates different keys for different number formats
        let key1 = XLSXEngine.formatToKey(currencyFormat)
        let key2 = XLSXEngine.formatToKey(percentageFormat)
        let key3 = XLSXEngine.formatToKey(customFormat)
        
        XCTAssertNotEqual(key1, key2, "Currency and percentage formats should have different keys")
        XCTAssertNotEqual(key1, key3, "Currency and custom formats should have different keys")
        XCTAssertNotEqual(key2, key3, "Percentage and custom formats should have different keys")
        
        // Verify that number format information is included in the key
        XCTAssertTrue(key1.contains("numberFormat:$#,##0.00"), "Currency format key should include number format")
        XCTAssertTrue(key2.contains("numberFormat:0.00%"), "Percentage format key should include number format")
        XCTAssertTrue(key3.contains("numberFormat:#,##0.00"), "Custom format key should include number format")
    }
    
    func testNumberFormatExcelGeneration() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Number Format Test")
        
        // Add cells with different number formats
        sheet.setCell("A1", number: 1234.56, format: .currency())
        sheet.setCell("A2", number: 0.1234, format: .percentage())
        
        var customFormat = CellFormat()
        customFormat.numberFormat = .custom
        customFormat.customNumberFormat = "$#,##0.00"
        sheet.setCell("A3", number: 5678.90, format: customFormat)
        
        // Save the workbook to test XLSX generation
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("number_format_test.xlsx")
        
        do {
            try workbook.save(to: tempURL)
            
            // Verify the file was created
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            
            // Clean up
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with number formats: \(error)")
        }
    }
    
    // MARK: - Text Wrapping Tests
    
    func testTextWrapping() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Text Wrapping Test")
        
        // Test text wrapping enabled
        var wrappedFormat = CellFormat()
        wrappedFormat.textWrapping = true
        wrappedFormat.verticalAlignment = .top
        
        sheet.setCell("A1", string: "This is a long text that should wrap to multiple lines when text wrapping is enabled.", format: wrappedFormat)
        
        // Test text wrapping disabled
        var noWrapFormat = CellFormat()
        noWrapFormat.textWrapping = false
        noWrapFormat.verticalAlignment = .top
        
        sheet.setCell("A2", string: "This text should not wrap and may be truncated.", format: noWrapFormat)
        
        // Test text wrapping with other formatting
        var wrappedBoldFormat = CellFormat()
        wrappedBoldFormat.textWrapping = true
        wrappedBoldFormat.verticalAlignment = .top
        wrappedBoldFormat.fontWeight = .bold
        wrappedBoldFormat.backgroundColor = "#E6E6E6"
        
        sheet.setCell("A3", string: "Bold wrapped text with background color.", format: wrappedBoldFormat)
        
        // Set column width to demonstrate wrapping
        sheet.setColumnWidth(1, width: 200)
        
        // Save and verify
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("text_wrapping_test.xlsx")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with text wrapping: \(error)")
        }
    }
    
    func testTextWrappingInFormatToKey() {
        var format1 = CellFormat()
        format1.textWrapping = true
        
        var format2 = CellFormat()
        format2.textWrapping = false
        
        let format3 = CellFormat() // nil textWrapping
        
        let key1 = XLSXEngine.formatToKey(format1)
        let key2 = XLSXEngine.formatToKey(format2)
        let key3 = XLSXEngine.formatToKey(format3)
        
        XCTAssertTrue(key1.contains("textWrapping:true"), "Text wrapping true should be in format key")
        XCTAssertTrue(key2.contains("textWrapping:false"), "Text wrapping false should be in format key")
        XCTAssertTrue(key3.contains("textWrapping:false"), "Nil text wrapping should default to false in format key")
    }
    
    // MARK: - Border Tests
    
    func testBordersActuallyWork() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border Test")
        
        // Test different border styles
        var thinBorderFormat = CellFormat.bordered()
        thinBorderFormat.borderTop = .thin
        thinBorderFormat.borderBottom = .thin
        thinBorderFormat.borderLeft = .thin
        thinBorderFormat.borderRight = .thin
        
        var mediumBorderFormat = CellFormat.bordered()
        mediumBorderFormat.borderTop = .medium
        mediumBorderFormat.borderBottom = .medium
        mediumBorderFormat.borderColor = "#FF0000" // Red borders
        
        var thickBorderFormat = CellFormat.bordered()
        thickBorderFormat.borderTop = .thick
        thickBorderFormat.borderBottom = .thick
        thickBorderFormat.borderColor = "#0000FF" // Blue borders
        
        // Set cells with different border styles
        sheet.setCell("A1", string: "Thin Borders", format: thinBorderFormat)
        sheet.setCell("A2", string: "Medium Red Borders", format: mediumBorderFormat)
        sheet.setCell("A3", string: "Thick Blue Borders", format: thickBorderFormat)
        
        // Verify border formats are stored correctly
        let thinCell = sheet.getCellWithFormat("A1")
        let mediumCell = sheet.getCellWithFormat("A2")
        let thickCell = sheet.getCellWithFormat("A3")
        
        XCTAssertNotNil(thinCell)
        XCTAssertNotNil(mediumCell)
        XCTAssertNotNil(thickCell)
        
        XCTAssertEqual(thinCell?.format?.borderTop, .thin)
        XCTAssertEqual(thinCell?.format?.borderBottom, .thin)
        XCTAssertEqual(thinCell?.format?.borderLeft, .thin)
        XCTAssertEqual(thinCell?.format?.borderRight, .thin)
        
        XCTAssertEqual(mediumCell?.format?.borderTop, .medium)
        XCTAssertEqual(mediumCell?.format?.borderBottom, .medium)
        XCTAssertEqual(mediumCell?.format?.borderColor, "#FF0000")
        
        XCTAssertEqual(thickCell?.format?.borderTop, .thick)
        XCTAssertEqual(thickCell?.format?.borderBottom, .thick)
        XCTAssertEqual(thickCell?.format?.borderColor, "#0000FF")
        
        // Test that format keys include border information
        if let thinFormat = thinCell?.format {
            let formatKey = XLSXEngine.formatToKey(thinFormat)
            XCTAssertTrue(formatKey.contains("borderTop:thin"))
            XCTAssertTrue(formatKey.contains("borderBottom:thin"))
            XCTAssertTrue(formatKey.contains("borderLeft:thin"))
            XCTAssertTrue(formatKey.contains("borderRight:thin"))
        }
        
        if let mediumFormat = mediumCell?.format {
            let formatKey = XLSXEngine.formatToKey(mediumFormat)
            XCTAssertTrue(formatKey.contains("borderTop:medium"))
            XCTAssertTrue(formatKey.contains("borderColor:#FF0000"))
        }
        
        // Test Excel file generation with borders
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("border_test.xlsx")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with borders: \(error)")
        }
    }
    
    func testDifferentBorderStyles() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border Styles Test")
        
        // Test all border styles
        let borderStyles: [BorderStyle] = [.thin, .medium, .thick]
        let borderColors = ["#FF0000", "#00FF00", "#0000FF"]
        
        for (index, style) in borderStyles.enumerated() {
            var borderFormat = CellFormat.bordered()
            borderFormat.borderTop = style
            borderFormat.borderBottom = style
            borderFormat.borderLeft = style
            borderFormat.borderRight = style
            borderFormat.borderColor = borderColors[index]
            
            let cellAddress = "A\(index + 1)"
            sheet.setCell(cellAddress, string: "Style: \(style.rawValue)", format: borderFormat)
            
            // Verify border style is stored correctly
            let cell = sheet.getCellWithFormat(cellAddress)
            XCTAssertNotNil(cell)
            XCTAssertEqual(cell?.format?.borderTop, style)
            XCTAssertEqual(cell?.format?.borderBottom, style)
            XCTAssertEqual(cell?.format?.borderLeft, style)
            XCTAssertEqual(cell?.format?.borderRight, style)
            XCTAssertEqual(cell?.format?.borderColor, borderColors[index])
        }
    }
    
    func testBorderWithOtherFormatting() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border with Formatting Test")
        
        // Test borders combined with other formatting
        var formattedBorderFormat = CellFormat.bordered()
        formattedBorderFormat.borderTop = .thick
        formattedBorderFormat.borderBottom = .thick
        formattedBorderFormat.borderColor = "#FF0000"
        formattedBorderFormat.fontWeight = .bold
        formattedBorderFormat.fontSize = 14.0
        formattedBorderFormat.fontColor = "#0000FF"
        formattedBorderFormat.backgroundColor = "#FFFF00"
        formattedBorderFormat.horizontalAlignment = .center
        formattedBorderFormat.verticalAlignment = .center
        
        sheet.setCell("A1", string: "Formatted with Borders", format: formattedBorderFormat)
        
        // Verify all formatting is applied
        let cell = sheet.getCellWithFormat("A1")
        XCTAssertNotNil(cell)
        
        XCTAssertEqual(cell?.format?.borderTop, .thick)
        XCTAssertEqual(cell?.format?.borderBottom, .thick)
        XCTAssertEqual(cell?.format?.borderColor, "#FF0000")
        XCTAssertEqual(cell?.format?.fontWeight, .bold)
        XCTAssertEqual(cell?.format?.fontSize, 14.0)
        XCTAssertEqual(cell?.format?.fontColor, "#0000FF")
        XCTAssertEqual(cell?.format?.backgroundColor, "#FFFF00")
        XCTAssertEqual(cell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(cell?.format?.verticalAlignment, .center)
        
        // Test that format key includes all formatting
        if let format = cell?.format {
            let formatKey = XLSXEngine.formatToKey(format)
            XCTAssertTrue(formatKey.contains("borderTop:thick"))
            XCTAssertTrue(formatKey.contains("borderColor:#FF0000"))
            XCTAssertTrue(formatKey.contains("fontWeight:bold"))
            XCTAssertTrue(formatKey.contains("fontColor:#0000FF"))
            XCTAssertTrue(formatKey.contains("backgroundColor:#FFFF00"))
            XCTAssertTrue(formatKey.contains("horizontalAlignment:center"))
            XCTAssertTrue(formatKey.contains("verticalAlignment:center"))
        }
    }
    
    // MARK: - Merge Tests
    
    func testMergedCellsActuallyWork() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Merge Test")
        
        // Test complex merging scenarios
        sheet.setCell("A1", string: "Merged A1:C1")
        sheet.setCell("A2", string: "Merged A2:B2")
        sheet.setCell("C2", string: "Single Cell")
        
        // Perform merges
        sheet.mergeCells("A1:C1")
        sheet.mergeCells("A2:B2")
        
        // Verify merged ranges are stored correctly
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, 2)
        
        // Check first merge (A1:C1)
        XCTAssertTrue(mergedRanges.contains { $0.excelRange == "A1:C1" })
        let firstMerge = mergedRanges.first { $0.excelRange == "A1:C1" }
        XCTAssertNotNil(firstMerge)
        XCTAssertEqual(firstMerge?.start.excelAddress, "A1")
        XCTAssertEqual(firstMerge?.end.excelAddress, "C1")
        
        // Check second merge (A2:B2)
        XCTAssertTrue(mergedRanges.contains { $0.excelRange == "A2:B2" })
        let secondMerge = mergedRanges.first { $0.excelRange == "A2:B2" }
        XCTAssertNotNil(secondMerge)
        XCTAssertEqual(secondMerge?.start.excelAddress, "A2")
        XCTAssertEqual(secondMerge?.end.excelAddress, "B2")
        
        // Verify cell values are preserved
        XCTAssertEqual(sheet.getCell("A1"), .string("Merged A1:C1"))
        XCTAssertEqual(sheet.getCell("A2"), .string("Merged A2:B2"))
        XCTAssertEqual(sheet.getCell("C2"), .string("Single Cell"))
        
        // Test Excel file generation with merged cells
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("merge_test.xlsx")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with merged cells: \(error)")
        }
    }
    
    func testComplexMerging() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Complex Merge Test")
        
        // Test various merge scenarios
        let mergeRanges = [
            "A1:B1",   // Horizontal merge
            "A2:A4",   // Vertical merge
            "C1:D3",   // 2x3 merge
            "F1:F1",   // Single cell (no merge)
            "G1:H2"    // 2x2 merge
        ]
        
        // Set data and perform merges
        for (index, range) in mergeRanges.enumerated() {
            let startCell = range.components(separatedBy: ":")[0]
            sheet.setCell(startCell, string: "Merge \(index + 1): \(range)")
            sheet.mergeCells(range)
        }
        
        // Verify all merges are stored correctly
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, mergeRanges.count)
        
        for range in mergeRanges {
            XCTAssertTrue(mergedRanges.contains { $0.excelRange == range })
        }
        
        // Test specific merge validations
        let horizontalMerge = mergedRanges.first { $0.excelRange == "A1:B1" }
        XCTAssertNotNil(horizontalMerge)
        XCTAssertEqual(horizontalMerge?.start.row, 1)
        XCTAssertEqual(horizontalMerge?.start.column, 1)
        XCTAssertEqual(horizontalMerge?.end.row, 1)
        XCTAssertEqual(horizontalMerge?.end.column, 2)
        
        let verticalMerge = mergedRanges.first { $0.excelRange == "A2:A4" }
        XCTAssertNotNil(verticalMerge)
        XCTAssertEqual(verticalMerge?.start.row, 2)
        XCTAssertEqual(verticalMerge?.start.column, 1)
        XCTAssertEqual(verticalMerge?.end.row, 4)
        XCTAssertEqual(verticalMerge?.end.column, 1)
        
        let largeMerge = mergedRanges.first { $0.excelRange == "C1:D3" }
        XCTAssertNotNil(largeMerge)
        XCTAssertEqual(largeMerge?.start.row, 1)
        XCTAssertEqual(largeMerge?.start.column, 3)
        XCTAssertEqual(largeMerge?.end.row, 3)
        XCTAssertEqual(largeMerge?.end.column, 4)
    }
    
    func testComplexBorderAndMergeCombination() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Border and Merge Test")
        
        // Test the exact scenario from the user's report
        var borderedFormat = CellFormat.bordered()
        borderedFormat.fontSize = 11
        borderedFormat.fontWeight = .bold
        borderedFormat.horizontalAlignment = .center
        borderedFormat.verticalAlignment = .center
        borderedFormat.fontName = "Calibri"
        borderedFormat.borderTop = .thin
        borderedFormat.borderBottom = .thin
        borderedFormat.borderLeft = .thin
        borderedFormat.borderRight = .thin
        borderedFormat.borderColor = "#000000"
        
        // Set cells with borders and formatting
        sheet.setCell("A1", string: "Test1", format: borderedFormat)
        sheet.setCell("A2", string: "Test2")
        
        // Perform merges
        sheet.mergeCells("A1:B1")
        sheet.mergeCells("A2:B2")
        
        // Set column widths
        sheet.setColumnWidth(1, width: 100.0)
        sheet.setColumnWidth(2, width: 150.0)
        
        // Verify everything is stored correctly
        let mergedRanges = sheet.getMergedRanges()
        XCTAssertEqual(mergedRanges.count, 2)
        XCTAssertTrue(mergedRanges.contains { $0.excelRange == "A1:B1" })
        XCTAssertTrue(mergedRanges.contains { $0.excelRange == "A2:B2" })
        
        // Verify border formatting
        let borderedCell = sheet.getCellWithFormat("A1")
        XCTAssertNotNil(borderedCell)
        XCTAssertEqual(borderedCell?.format?.fontSize, 11)
        XCTAssertEqual(borderedCell?.format?.fontWeight, .bold)
        XCTAssertEqual(borderedCell?.format?.horizontalAlignment, .center)
        XCTAssertEqual(borderedCell?.format?.verticalAlignment, .center)
        XCTAssertEqual(borderedCell?.format?.fontName, "Calibri")
        XCTAssertEqual(borderedCell?.format?.borderTop, .thin)
        XCTAssertEqual(borderedCell?.format?.borderBottom, .thin)
        XCTAssertEqual(borderedCell?.format?.borderLeft, .thin)
        XCTAssertEqual(borderedCell?.format?.borderRight, .thin)
        XCTAssertEqual(borderedCell?.format?.borderColor, "#000000")
        
        // Verify column widths
        XCTAssertEqual(sheet.getColumnWidth(1), 100.0)
        XCTAssertEqual(sheet.getColumnWidth(2), 150.0)
        
        // Test Excel file generation
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("border_merge_test.xlsx")
        
        do {
            try workbook.save(to: tempURL)
            XCTAssertTrue(FileManager.default.fileExists(atPath: tempURL.path))
            try FileManager.default.removeItem(at: tempURL)
        } catch {
            XCTFail("Failed to save workbook with borders and merges: \(error)")
        }
    }
} 