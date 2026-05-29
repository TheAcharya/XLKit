//
//  SheetProtectionTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
@testable import XLKit
@testable import XLKitXLSX

@MainActor
final class SheetProtectionTests: XLKitTestBase {
    
    func testDefaultProtectionIsNil() {
        let sheet = Workbook().addSheet(name: "Sheet1")
        XCTAssertNil(sheet.protection)
    }
    
    func testDefaultProtectionStructEnablesSheet() {
        let protection = SheetProtection()
        XCTAssertEqual(protection.sheet, true)
        XCTAssertNil(protection.formatCells)
        XCTAssertNil(protection.password)
    }
    
    func testMinimalProtectionXML() {
        let protection = SheetProtection()
        XCTAssertEqual(XLSXEngine.sheetProtectionXML(protection), "<sheetProtection sheet=\"1\"/>")
    }
    
    func testProtectionWithLegacyPassword() {
        var protection = SheetProtection()
        protection.password = "CC3F"
        protection.selectLockedCells = true
        XCTAssertEqual(
            XLSXEngine.sheetProtectionXML(protection),
            "<sheetProtection sheet=\"1\" password=\"CC3F\" selectLockedCells=\"1\"/>"
        )
    }
    
    func testProtectionWithModernHash() {
        var protection = SheetProtection()
        protection.algorithmName = "SHA-512"
        protection.hashValue = "abc=="
        protection.saltValue = "def=="
        protection.spinCount = 100_000
        let xml = XLSXEngine.sheetProtectionXML(protection)
        XCTAssertTrue(xml.contains("algorithmName=\"SHA-512\""))
        XCTAssertTrue(xml.contains("hashValue=\"abc==\""))
        XCTAssertTrue(xml.contains("saltValue=\"def==\""))
        XCTAssertTrue(xml.contains("spinCount=\"100000\""))
    }
    
    func testProtectionWithGranularPermissions() {
        var protection = SheetProtection()
        protection.formatCells = false
        protection.insertRows = false
        protection.sort = true
        let xml = XLSXEngine.sheetProtectionXML(protection)
        XCTAssertTrue(xml.contains("formatCells=\"0\""))
        XCTAssertTrue(xml.contains("insertRows=\"0\""))
        XCTAssertTrue(xml.contains("sort=\"1\""))
    }
    
    func testProtectionOmittedWhenNotConfigured() {
        let sheet = Workbook().addSheet(name: "Sheet1")
        XCTAssertNil(sheet.protection)
    }
    
    func testProtectionAllAttributesEmitted() {
        var protection = SheetProtection()
        protection.password = "CC3F"
        protection.algorithmName = "SHA-512"
        protection.hashValue = "h"
        protection.saltValue = "s"
        protection.spinCount = 1
        protection.objects = true
        protection.scenarios = true
        protection.formatCells = true
        protection.formatColumns = true
        protection.formatRows = true
        protection.insertColumns = true
        protection.insertRows = true
        protection.insertHyperlinks = true
        protection.deleteColumns = true
        protection.deleteRows = true
        protection.selectLockedCells = true
        protection.selectUnlockedCells = true
        protection.sort = true
        protection.autoFilter = true
        protection.pivotTables = true
        
        let xml = XLSXEngine.sheetProtectionXML(protection)
        let expectedAttributes = [
            "sheet", "password", "algorithmName", "hashValue", "saltValue", "spinCount",
            "objects", "scenarios",
            "formatCells", "formatColumns", "formatRows",
            "insertColumns", "insertRows", "insertHyperlinks",
            "deleteColumns", "deleteRows",
            "selectLockedCells", "selectUnlockedCells",
            "sort", "autoFilter", "pivotTables",
        ]
        for name in expectedAttributes {
            XCTAssertTrue(xml.contains("\(name)=\""), "Missing attribute \(name) in: \(xml)")
        }
    }
    
    func testSavingWorkbookWithProtectedSheetSucceeds() throws {
        let workbook = Workbook()
        _ = workbook.addSheet(name: "Main")
        let tech = workbook.addSheet(name: "Strings")
        tech.protection = SheetProtection()
        let url = makeTempWorkbookURL(prefix: "protected_sheet")
        defer { cleanupTempFile(at: url) }
        try workbook.save(to: url)
        XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
    }
}
