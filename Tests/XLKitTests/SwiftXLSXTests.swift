import XCTest
@testable import XLKit

@available(macOS 13.0, *)
final class XLKitTests: XCTestCase {
    
    func testCreateWorkbook() throws {
        let workbook = XLKit.createWorkbook()
        XCTAssertNotNil(workbook)
        XCTAssertEqual(workbook.sheetCount, 0)
    }
    
    func testAddSheet() throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test Sheet")
        
        XCTAssertEqual(workbook.sheetCount, 1)
        XCTAssertEqual(sheet.name, "Test Sheet")
        XCTAssertEqual(sheet.cellCount, 0)
    }
    
    func testAddCells() throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test Sheet")
        
        // Add different types of cells
        let textCell = sheet.addCell(coordinate: CellCoordinate(row: 1, column: 1), text: "Hello, World!")
        let numberCell = sheet.addCell(coordinate: CellCoordinate(row: 1, column: 2), number: 42.5)
        let integerCell = sheet.addCell(coordinate: CellCoordinate(row: 1, column: 3), integer: 100)
        let booleanCell = sheet.addCell(coordinate: CellCoordinate(row: 1, column: 4), boolean: true)
        let dateCell = sheet.addCell(coordinate: CellCoordinate(row: 1, column: 5), date: Date())
        let formulaCell = sheet.addCell(coordinate: CellCoordinate(row: 2, column: 1), formula: "=SUM(B1:C1)")
        
        XCTAssertEqual(sheet.cellCount, 6)
        XCTAssertEqual(textCell.value, .text("Hello, World!"))
        XCTAssertEqual(numberCell.value, .number(42.5))
        XCTAssertEqual(integerCell.value, .integer(100))
        XCTAssertEqual(booleanCell.value, .boolean(true))
        XCTAssertEqual(formulaCell.value, .formula("=SUM(B1:C1)"))
    }
    
    func testCellCoordinate() throws {
        let coord1 = CellCoordinate(row: 1, column: 1)
        XCTAssertEqual(coord1.address, "A1")
        
        let coord2 = CellCoordinate(row: 5, column: 26)
        XCTAssertEqual(coord2.address, "Z5")
        
        let coord3 = CellCoordinate(row: 10, column: 27)
        XCTAssertEqual(coord3.address, "AA10")
        
        // Test parsing from address
        let parsed1 = CellCoordinate(address: "B3")
        XCTAssertNotNil(parsed1)
        XCTAssertEqual(parsed1?.row, 3)
        XCTAssertEqual(parsed1?.column, 2)
        
        let parsed2 = CellCoordinate(address: "AA15")
        XCTAssertNotNil(parsed2)
        XCTAssertEqual(parsed2?.row, 15)
        XCTAssertEqual(parsed2?.column, 27)
    }
    
    func testCellRange() throws {
        let range = CellRange(startRow: 1, startColumn: 1, endRow: 5, endColumn: 3)
        XCTAssertEqual(range.address, "A1:C5")
        
        let coord1 = CellCoordinate(row: 3, column: 2)
        XCTAssertTrue(range.contains(coord1))
        
        let coord2 = CellCoordinate(row: 6, column: 2)
        XCTAssertFalse(range.contains(coord2))
    }
    
    func testFontAndColor() throws {
        let font = Font(standardName: .calibri, size: 12, isBold: true, color: .blue)
        XCTAssertEqual(font.name, "Calibri")
        XCTAssertEqual(font.size, 12)
        XCTAssertTrue(font.isBold)
        XCTAssertEqual(font.color, .blue)
        
        let color = Color(red: 255, green: 0, blue: 0)
        XCTAssertEqual(color.hexString, "FF0000")
        
        let colorFromHex = Color(hex: "#00FF00")
        XCTAssertNotNil(colorFromHex)
        XCTAssertEqual(colorFromHex?.hexString, "00FF00")
    }
    
    func testImageManager() throws {
        let imageManager = ImageManager.shared
        
        // Create a simple test image
        let testImage = NSImage(size: CGSize(width: 100, height: 100))
        testImage.lockFocus()
        NSColor.red.setFill()
        NSRect(origin: .zero, size: testImage.size).fill()
        testImage.unlockFocus()
        
        let imageId = try imageManager.addImage(testImage)
        XCTAssertFalse(imageId.isEmpty)
        
        let retrievedImage = imageManager.getImage(id: imageId)
        XCTAssertNotNil(retrievedImage)
        XCTAssertEqual(retrievedImage?.id, imageId)
    }
    
    func testSheetOperations() throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test Sheet")
        
        // Add data in rows and columns
        let rowData: [CellValue] = [.text("Name"), .text("Age"), .text("City")]
        sheet.addRow(1, data: rowData)
        
        let columnData: [CellValue] = [.text("John"), .integer(25), .text("New York")]
        sheet.addColumn(1, data: columnData, startRow: 2)
        
        XCTAssertEqual(sheet.cellCount, 6)
        
        // Test getting row and column data
        let retrievedRow = sheet.getRowData(1)
        XCTAssertEqual(retrievedRow.count, 3)
        XCTAssertEqual(retrievedRow[0], .text("Name"))
        
        let retrievedColumn = sheet.getColumnData(1)
        XCTAssertEqual(retrievedColumn.count, 3)
        XCTAssertEqual(retrievedColumn[0], .text("John"))
    }
    
    func testMergedCells() throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test Sheet")
        
        let range = CellRange(startRow: 1, startColumn: 1, endRow: 2, endColumn: 3)
        sheet.mergeCells(in: range)
        
        XCTAssertEqual(sheet.allMergedRanges.count, 1)
        XCTAssertEqual(sheet.allMergedRanges[0], range)
        
        let coord = CellCoordinate(row: 1, column: 2)
        let mergedRange = sheet.getMergedRange(for: coord)
        XCTAssertNotNil(mergedRange)
        XCTAssertEqual(mergedRange, range)
    }
    
    func testColumnAndRowManagement() throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Test Sheet")
        
        sheet.setColumnWidth(1, width: 15.0)
        sheet.setRowHeight(1, height: 20.0)
        
        XCTAssertEqual(sheet.getColumnWidth(1), 15.0)
        XCTAssertEqual(sheet.getRowHeight(1), 20.0)
        
        // Test auto-fit
        sheet.addCell(coordinate: CellCoordinate(row: 1, column: 1), text: "Very long text for testing")
        sheet.autoFitColumn(1)
        
        let newWidth = sheet.getColumnWidth(1)
        XCTAssertGreaterThan(newWidth, 15.0) // Should be wider after auto-fit
    }
    
    func testWorkbookProperties() throws {
        let properties = WorkbookProperties(
            title: "Test Workbook",
            author: "Test Author",
            company: "Test Company"
        )
        
        let workbook = Workbook(properties: properties)
        XCTAssertEqual(workbook.properties.title, "Test Workbook")
        XCTAssertEqual(workbook.properties.author, "Test Author")
        XCTAssertEqual(workbook.properties.company, "Test Company")
    }
    
    func testUtilityFunctions() throws {
        // Test checksum
        let data = "Hello, World!".data(using: .utf8)!
        let checksum = XLKitUtils.checksum(data: data)
        XCTAssertFalse(checksum.isEmpty)
        XCTAssertEqual(checksum.count, 64) // SHA256 hex string length
        
        // Test cell reference parsing
        let parsed = XLKitUtils.parseCellReference("B3")
        XCTAssertNotNil(parsed)
        XCTAssertEqual(parsed?.column, 2)
        XCTAssertEqual(parsed?.row, 3)
        
        // Test cell reference creation
        let reference = XLKitUtils.createCellReference(column: 5, row: 10)
        XCTAssertEqual(reference, "E10")
        
        // Test column letter conversion
        XCTAssertEqual(XLKitUtils.columnLetterToNumber("A"), 1)
        XCTAssertEqual(XLKitUtils.columnLetterToNumber("Z"), 26)
        XCTAssertEqual(XLKitUtils.columnLetterToNumber("AA"), 27)
        
        XCTAssertEqual(XLKitUtils.columnNumberToLetter(1), "A")
        XCTAssertEqual(XLKitUtils.columnNumberToLetter(26), "Z")
        XCTAssertEqual(XLKitUtils.columnNumberToLetter(27), "AA")
        
        // Test XML escaping
        let escaped = XLKitUtils.escapeXML("<test>&\"'</test>")
        XCTAssertEqual(escaped, "&lt;test&gt;&amp;&quot;&#x27;&lt;/test&gt;")
        
        // Test sheet name sanitization
        let sanitized = XLKitUtils.sanitizeSheetName("Test/Sheet*Name?")
        XCTAssertEqual(sanitized, "Test Sheet Name")
    }
    
    func testGenerateTestWorkbook() throws {
        let workbook = XLKit.generateTestWorkbook()
        XCTAssertNotNil(workbook)
        XCTAssertGreaterThan(workbook.sheetCount, 0)
        
        let sheet = workbook.getSheet(at: 0)
        XCTAssertNotNil(sheet)
        XCTAssertGreaterThan(sheet?.cellCount ?? 0, 0)
    }
    
    func testErrorHandling() throws {
        // Test invalid cell coordinate
        XCTAssertThrowsError(try CellCoordinate(row: 0, column: 1)) { error in
            XCTAssertTrue(error is FatalError)
        }
        
        XCTAssertThrowsError(try CellCoordinate(row: 1, column: 0)) { error in
            XCTAssertTrue(error is FatalError)
        }
    }
    
    func testPerformance() throws {
        let workbook = XLKit.createWorkbook()
        let sheet = workbook.addSheet(name: "Performance Test")
        
        measure {
            // Add 1000 cells
            for row in 1...100 {
                for col in 1...10 {
                    let value = row * col
                    sheet.addCell(coordinate: CellCoordinate(row: row, column: col), integer: value)
                }
            }
        }
        
        XCTAssertEqual(sheet.cellCount, 1000)
    }
}
