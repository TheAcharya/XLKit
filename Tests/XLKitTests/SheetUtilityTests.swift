//
//  SheetUtilityTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class SheetUtilityTests: XLKitTestBase {
    
    func testColumnAndRowSizing() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setColumnWidth(1, width: 100.0)
        sheet.setRowHeight(1, height: 25.0)
        
        XCTAssertEqual(sheet.getColumnWidth(1), 100.0)
        XCTAssertEqual(sheet.getRowHeight(1), 25.0)
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
        sheet.setCell("E1", date: Self.fixedTestDate)
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
