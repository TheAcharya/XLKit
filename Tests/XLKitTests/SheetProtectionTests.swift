//
//  SheetProtectionTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
@testable import XLKit
@testable import XLKitXLSX

/// Must match `ComprehensiveDemoProtection` in XLKitTestRunner (reproducible comprehensive demo).
private enum ComprehensiveDemoProtectionFixtures {
    static let password = "1234"
    static let passwordSheetSalt = Data("XLKitPassDemo1!!".utf8)
    static let modernSheetSalt = Data("XLKitModHashDemo!".utf8)
}

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
    
    func testLegacyPasswordHashFor1234() {
        XCTAssertEqual(CoreUtils.excelLegacySheetPasswordHash(for: "1234"), "CC3D")
    }
    
    func testModernPasswordHashFor1234ComprehensiveDemoPasswordSheet() throws {
        let modern = try CoreUtils.excelModernSheetPasswordHash(
            for: "1234",
            salt: ComprehensiveDemoProtectionFixtures.passwordSheetSalt
        )
        XCTAssertEqual(modern.algorithmName, "SHA-512")
        XCTAssertEqual(modern.saltValue, "WExLaXRQYXNzRGVtbzEhIQ==")
        XCTAssertEqual(
            modern.hashValue,
            "r5j81/AZyLlOvDcoZdtgz0KZV/LU28sGe5LjE5dbkmalhvrdLKmD7xSLn403UyjjZwyOqjApNK3taOIZ8qZIYQ=="
        )
        XCTAssertEqual(modern.spinCount, 100_000)
    }
    
    func testModernPasswordHashFor1234ComprehensiveDemoModernSheet() throws {
        let modern = try CoreUtils.excelModernSheetPasswordHash(
            for: "1234",
            salt: ComprehensiveDemoProtectionFixtures.modernSheetSalt
        )
        XCTAssertEqual(modern.saltValue, "WExLaXRNb2RIYXNoRGVtbyE=")
        XCTAssertEqual(
            modern.hashValue,
            "UUuu11+Bm0SBR7dA+v1Fl7xnQTS/EYuHV2QIjv45LWN3a3xxWaIDiu2/VdJz4h2nivwAeU5N8oZBe5ClFu/IaQ=="
        )
    }
    
    func testConfigureSheetPasswordAppliesLegacyAndModern() throws {
        var protection = SheetProtection()
        try CoreUtils.configureSheetPassword(
            &protection,
            plaintext: "789648",
            salt: Data("XLKitTestSalt!!".utf8)
        )
        XCTAssertEqual(protection.password, CoreUtils.excelLegacySheetPasswordHash(for: "789648"))
        XCTAssertEqual(protection.algorithmName, "SHA-512")
        XCTAssertNotNil(protection.saltValue)
        XCTAssertNotNil(protection.hashValue)
        XCTAssertEqual(protection.spinCount, 100_000)
    }
    
    func testConfigureSheetPasswordRejectsEmptyPassword() {
        var protection = SheetProtection()
        XCTAssertThrowsError(try CoreUtils.configureSheetPassword(&protection, plaintext: "")) { error in
            guard case .securityError(let message) = error as? XLKitError else {
                return XCTFail("Expected XLKitError.securityError, got \(error)")
            }
            XCTAssertTrue(message.contains("must not be empty"))
        }
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
