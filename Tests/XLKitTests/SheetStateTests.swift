//
//  SheetStateTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
@testable import XLKit
@testable import XLKitXLSX

@Suite
@MainActor
struct SheetStateTests {
    
    @Test func testDefaultStateIsVisible() {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Sheet1")
        #expect(sheet.state == .visible)
    }
    
    @Test func testVisibleSheetOmitsStateAttribute() {
        let sheet = Workbook().addSheet(name: "Sheet1")
        #expect(XLSXEngine.sheetStateAttribute(sheet) == "")
    }
    
    @Test func testHiddenSheetEmitsStateAttribute() {
        let sheet = Workbook().addSheet(name: "Strings")
        sheet.state = .hidden
        #expect(XLSXEngine.sheetStateAttribute(sheet) == " state=\"hidden\"")
    }
    
    @Test func testVeryHiddenSheetEmitsStateAttribute() {
        let sheet = Workbook().addSheet(name: "Secret")
        sheet.state = .veryHidden
        #expect(XLSXEngine.sheetStateAttribute(sheet) == " state=\"veryHidden\"")
    }
    
    @Test func testActiveTabOmittedWhenFirstSheetVisible() {
        let workbook = Workbook()
        _ = workbook.addSheet(name: "Main")
        let tech = workbook.addSheet(name: "Strings")
        tech.state = .hidden
        #expect(XLSXEngine.activeTabAttribute(for: workbook.getSheets()) == "")
    }
    
    @Test func testActiveTabPointsAtFirstVisibleSheet() {
        let workbook = Workbook()
        let first = workbook.addSheet(name: "Hidden1")
        let second = workbook.addSheet(name: "Hidden2")
        _ = workbook.addSheet(name: "Visible")
        first.state = .hidden
        second.state = .hidden
        #expect(XLSXEngine.activeTabAttribute(for: workbook.getSheets()) == " activeTab=\"2\"")
    }
    
    @Test func testSavingWorkbookWithHiddenSheetSucceeds() throws {
        let workbook = Workbook()
        _ = workbook.addSheet(name: "Main")
        let tech = workbook.addSheet(name: "Strings")
        tech.state = .hidden
        let url = XLKitTestSupport.makeTempWorkbookURL(prefix: "hidden_state")
        defer { XLKitTestSupport.cleanupTempFile(at: url) }
        try workbook.save(to: url)
        #expect(FileManager.default.fileExists(atPath: url.path))
    }
}
