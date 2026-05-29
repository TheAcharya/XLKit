//
//  SheetStateTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
@testable import XLKit
@testable import XLKitXLSX

@MainActor
final class SheetStateTests: XLKitTestBase {
    
    func testDefaultStateIsVisible() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Sheet1")
        XCTAssertEqual(sheet.state, .visible)
    }
    
    func testVisibleSheetOmitsStateAttribute() {
        let sheet = Workbook().addSheet(name: "Sheet1")
        XCTAssertEqual(XLSXEngine.sheetStateAttribute(sheet), "")
    }
    
    func testHiddenSheetEmitsStateAttribute() {
        let sheet = Workbook().addSheet(name: "Strings")
        sheet.state = .hidden
        XCTAssertEqual(XLSXEngine.sheetStateAttribute(sheet), " state=\"hidden\"")
    }
    
    func testVeryHiddenSheetEmitsStateAttribute() {
        let sheet = Workbook().addSheet(name: "Secret")
        sheet.state = .veryHidden
        XCTAssertEqual(XLSXEngine.sheetStateAttribute(sheet), " state=\"veryHidden\"")
    }
    
    func testActiveTabOmittedWhenFirstSheetVisible() {
        let workbook = Workbook()
        _ = workbook.addSheet(name: "Main")
        let tech = workbook.addSheet(name: "Strings")
        tech.state = .hidden
        XCTAssertEqual(XLSXEngine.activeTabAttribute(for: workbook.getSheets()), "")
    }
    
    func testActiveTabPointsAtFirstVisibleSheet() {
        let workbook = Workbook()
        let first = workbook.addSheet(name: "Hidden1")
        let second = workbook.addSheet(name: "Hidden2")
        _ = workbook.addSheet(name: "Visible")
        first.state = .hidden
        second.state = .hidden
        XCTAssertEqual(XLSXEngine.activeTabAttribute(for: workbook.getSheets()), " activeTab=\"2\"")
    }
    
    func testSavingWorkbookWithHiddenSheetSucceeds() throws {
        let workbook = Workbook()
        _ = workbook.addSheet(name: "Main")
        let tech = workbook.addSheet(name: "Strings")
        tech.state = .hidden
        let url = makeTempWorkbookURL(prefix: "hidden_state")
        defer { cleanupTempFile(at: url) }
        try workbook.save(to: url)
        XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
    }
}
