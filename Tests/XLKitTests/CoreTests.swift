//
//  CoreTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct CoreTests {
    
    @Test func testCreateWorkbook() throws {
        let workbook = Workbook()
        #expect(workbook.getSheets().count == 0)
    }
    
    @Test func testAddSheet() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test Sheet")
        
        #expect(workbook.getSheets().count == 1)
        #expect(sheet.name == "Test Sheet")
        #expect(sheet.id == 1)
    }
    
    @Test func testGetSheetByName() throws {
        let workbook = Workbook()
        let sheet1 = workbook.addSheet(name: "Sheet1")
        let sheet2 = workbook.addSheet(name: "Sheet2")
        
        #expect(workbook.getSheet(name: "Sheet1") == sheet1)
        let retrievedSheet2 = workbook.getSheet(name: "Sheet2")
        #expect(retrievedSheet2 != nil)
        #expect(retrievedSheet2?.name == "Sheet2")
        #expect(retrievedSheet2?.id == sheet2.id)
        #expect(retrievedSheet2 == sheet2)
        #expect(workbook.getSheet(name: "NonExistent") == nil)
    }
    
    @Test func testRemoveSheet() throws {
        let workbook = Workbook()
        _ = workbook.addSheet(name: "Sheet1")
        _ = workbook.addSheet(name: "Sheet2")
        
        #expect(workbook.getSheets().count == 2)
        
        workbook.removeSheet(name: "Sheet1")
        #expect(workbook.getSheets().count == 1)
        #expect(workbook.getSheet(name: "Sheet1") == nil)
        #expect(workbook.getSheet(name: "Sheet2") != nil)
    }
    
    @Test func testClearSheet() throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        
        sheet.setCell("A1", value: .string("Test"))
        sheet.mergeCells("B1:C1")
        
        #expect(sheet.getUsedCells().count == 1)
        #expect(sheet.getMergedRanges().count == 1)
        
        sheet.clear()
        
        #expect(sheet.getUsedCells().count == 0)
        #expect(sheet.getMergedRanges().count == 0)
    }
}
