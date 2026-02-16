//
//  CoreTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class CoreTests: XLKitTestBase {
    
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
        let retrievedSheet2 = workbook.getSheet(name: "Sheet2")
        XCTAssertNotNil(retrievedSheet2)
        XCTAssertEqual(retrievedSheet2?.name, "Sheet2")
        XCTAssertEqual(retrievedSheet2?.id, sheet2.id)
        XCTAssertEqual(retrievedSheet2, sheet2)
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
}
