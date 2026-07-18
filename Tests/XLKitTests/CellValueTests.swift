//
//  CellValueTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct CellValueTests {
    
    @Test func testSetAndGetCell() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        // Test string value
        sheet.setCell("A1", value: .string("Hello"))
        #expect(sheet.getCell("A1") == .string("Hello"))
        
        // Test number value
        sheet.setCell("B1", value: .number(42.5))
        #expect(sheet.getCell("B1") == .number(42.5))
        
        // Test integer value
        sheet.setCell("C1", value: .integer(100))
        #expect(sheet.getCell("C1") == .integer(100))
        
        // Test boolean value
        sheet.setCell("D1", value: .boolean(true))
        #expect(sheet.getCell("D1") == .boolean(true))
        
        // Test date value using a fixed date for deterministic tests
        sheet.setCell("E1", value: .date(XLKitTestSupport.fixedTestDate))
        #expect(sheet.getCell("E1") == .date(XLKitTestSupport.fixedTestDate))
        
        // Test formula
        sheet.setCell("F1", value: .formula("=A1+B1"))
        #expect(sheet.getCell("F1") == .formula("=A1+B1"))
    }
    
    @Test func testSetCellByRowColumn() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell(row: 1, column: 1, value: .string("A1"))
        sheet.setCell(row: 2, column: 2, value: .string("B2"))
        
        #expect(sheet.getCell("A1") == .string("A1"))
        #expect(sheet.getCell("B2") == .string("B2"))
    }
    
    @Test func testSetRange() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setRange("A1:C3", string: "Range")
        
        #expect(sheet.getCell("A1") == .string("Range"))
        #expect(sheet.getCell("B2") == .string("Range"))
        #expect(sheet.getCell("C3") == .string("Range"))
    }
    
    @Test func testGetUsedCells() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell("A1", value: .string("A1"))
        sheet.setCell("B2", value: .string("B2"))
        sheet.setCell("C3", value: .string("C3"))
        
        let usedCells = sheet.getUsedCells()
        #expect(usedCells.count == 3)
        #expect(usedCells.contains("A1"))
        #expect(usedCells.contains("B2"))
        #expect(usedCells.contains("C3"))
    }
    
    @Test func testCellValueStringValue() {
        #expect(CellValue.string("Hello").stringValue == "Hello")
        #expect(CellValue.number(42.5).stringValue == "42.5")
        #expect(CellValue.integer(100).stringValue == "100")
        #expect(CellValue.boolean(true).stringValue == "TRUE")
        #expect(CellValue.boolean(false).stringValue == "FALSE")
        #expect(CellValue.formula("=A1+B1").stringValue == "=A1+B1")
        #expect(CellValue.empty.stringValue == "")
        
        // Verify that date values produce a non-empty, stable string representation.
        let dateString = CellValue.date(XLKitTestSupport.epochDate).stringValue
        #expect(!(dateString.isEmpty))
        #expect(dateString.contains("1970"), "Expected epoch date string to contain year 1970, got: \(dateString)")
    }
    
    @Test func testCellValueType() {
        #expect(CellValue.string("").type == "string")
        #expect(CellValue.number(0).type == "number")
        #expect(CellValue.integer(0).type == "integer")
        #expect(CellValue.boolean(true).type == "boolean")
        #expect(CellValue.date(XLKitTestSupport.epochDate).type == "date")
        #expect(CellValue.formula("=A1").type == "formula")
        #expect(CellValue.empty.type == "empty")
    }
}
