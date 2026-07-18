//
//  SheetUtilityTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct SheetUtilityTests {
    
    @Test func testColumnAndRowSizing() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setColumnWidth(1, width: 100.0)
        sheet.setRowHeight(1, height: 25.0)
        
        #expect(sheet.getColumnWidth(1) == 100.0)
        #expect(sheet.getRowHeight(1) == 25.0)
    }
    
    @Test func testFluentAPI() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test fluent API with method chaining
        sheet
            .setCell("A1", value: .string("Header"))
            .setCell("B1", value: .string("Value"))
            .setRow(2, strings: ["Row1", "Data1"])
            .setRow(3, strings: ["Row2", "Data2"])
            .mergeCells("A1:B1")
        
        #expect(sheet.getCell("A1") == .string("Header"))
        #expect(sheet.getCell("B1") == .string("Value"))
        #expect(sheet.getCell("A2") == .string("Row1"))
        #expect(sheet.getCell("B2") == .string("Data1"))
        #expect(sheet.getCell("A3") == .string("Row2"))
        #expect(sheet.getCell("B3") == .string("Data2"))
        
        let mergedRanges = sheet.getMergedRanges()
        #expect(mergedRanges.count == 1)
        #expect(mergedRanges[0].excelRange == "A1:B1")
    }
    
    @Test func testSheetConvenienceMethods() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test convenience methods
        sheet.setCell("A1", string: "String Value")
        sheet.setCell("B1", number: 123.45)
        sheet.setCell("C1", integer: 42)
        sheet.setCell("D1", boolean: true)
        sheet.setCell("E1", date: XLKitTestSupport.fixedTestDate)
        sheet.setCell("F1", formula: "=A1+B1")
        
        #expect(sheet.getCell("A1") == .string("String Value"))
        #expect(sheet.getCell("B1") == .number(123.45))
        #expect(sheet.getCell("C1") == .integer(42))
        #expect(sheet.getCell("D1") == .boolean(true))
        #expect(sheet.getCell("E1")?.type == "date")
        #expect(sheet.getCell("F1") == .formula("=A1+B1"))
    }
    
    @Test func testSheetRowAndColumnMethods() {
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
        
        #expect(sheet.getCell("A1") == .string("Name"))
        #expect(sheet.getCell("B1") == .string("Age"))
        #expect(sheet.getCell("C1") == .string("Salary"))
        #expect(sheet.getCell("A2") == .number(30.0))
        #expect(sheet.getCell("B2") == .number(25.0))
        #expect(sheet.getCell("C2") == .number(50000.0))
        #expect(sheet.getCell("A3") == .integer(1))
        #expect(sheet.getCell("B3") == .integer(2))
        #expect(sheet.getCell("C3") == .integer(3))
        
        #expect(sheet.getCell("D1") == .string("Col1"))
        #expect(sheet.getCell("D2") == .string("Col2"))
        #expect(sheet.getCell("D3") == .string("Col3"))
        #expect(sheet.getCell("E1") == .number(10.5))
        #expect(sheet.getCell("E2") == .number(20.5))
        #expect(sheet.getCell("E3") == .number(30.5))
        #expect(sheet.getCell("F1") == .integer(100))
        #expect(sheet.getCell("F2") == .integer(200))
        #expect(sheet.getCell("F3") == .integer(300))
    }
    
    @Test func testSheetUtilityProperties() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test empty sheet
        #expect(sheet.isEmpty)
        #expect(sheet.cellCount == 0)
        #expect(sheet.imageCount == 0)
        
        // Add some data
        sheet.setCell("A1", value: .string("Test"))
        sheet.setCell("B2", value: .number(42.5))
        
        #expect(!(sheet.isEmpty))
        #expect(sheet.cellCount == 2)
        
        // Test allCells property
        let allCells = sheet.allCells
        #expect(allCells.count == 2)
        #expect(allCells["A1"] == .string("Test"))
        #expect(allCells["B2"] == .number(42.5))
        
        // Test allFormattedCells property
        let allFormattedCells = sheet.allFormattedCells
        #expect(allFormattedCells.count == 2)
        #expect(allFormattedCells["A1"]?.value == .string("Test"))
        #expect(allFormattedCells["B2"]?.value == .number(42.5))
    }
    
    @Test func testSheetConvenienceInitializer() {
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
        
        #expect(sheet.getCell("A1") == CellValue.string("Header"))
        #expect(sheet.getCell("B1") == CellValue.string("Value"))
        #expect(sheet.getCell("A2") == CellValue.number(42.5))
        #expect(sheet.cellCount == 3)
    }
}
