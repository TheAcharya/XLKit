//
//  SheetProtectionTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
@testable import XLKit
@testable import XLKitXLSX

/// Must match `ComprehensiveDemoProtection` in XLKitTestRunner (reproducible comprehensive demo).
private enum ComprehensiveDemoProtectionFixtures {
    static let password = "1234"
    static let passwordSheetSalt = Data("XLKitPassDemo1!!".utf8)
    static let modernSheetSalt = Data("XLKitModHashDemo!".utf8)
}

@Suite
@MainActor
struct SheetProtectionTests {
    
    @Test func testDefaultProtectionIsNil() {
        let sheet = Workbook().addSheet(name: "Sheet1")
        #expect(sheet.protection == nil)
    }
    
    @Test func testDefaultProtectionStructEnablesSheet() {
        let protection = SheetProtection()
        #expect(protection.sheet == true)
        #expect(protection.formatCells == nil)
        #expect(protection.password == nil)
    }
    
    @Test func testLegacyPasswordHashFor1234() {
        #expect(CoreUtils.excelLegacySheetPasswordHash(for: "1234") == "CC3D")
    }
    
    @Test func testModernPasswordHashFor1234ComprehensiveDemoPasswordSheet() throws {
        let modern = try CoreUtils.excelModernSheetPasswordHash(
            for: "1234",
            salt: ComprehensiveDemoProtectionFixtures.passwordSheetSalt
        )
        #expect(modern.algorithmName == "SHA-512")
        #expect(modern.saltValue == "WExLaXRQYXNzRGVtbzEhIQ==")
        #expect(
            modern.hashValue == "r5j81/AZyLlOvDcoZdtgz0KZV/LU28sGe5LjE5dbkmalhvrdLKmD7xSLn403UyjjZwyOqjApNK3taOIZ8qZIYQ=="
        )
        #expect(modern.spinCount == 100_000)
    }
    
    @Test func testModernPasswordHashFor1234ComprehensiveDemoModernSheet() throws {
        let modern = try CoreUtils.excelModernSheetPasswordHash(
            for: "1234",
            salt: ComprehensiveDemoProtectionFixtures.modernSheetSalt
        )
        #expect(modern.saltValue == "WExLaXRNb2RIYXNoRGVtbyE=")
        #expect(
            modern.hashValue == "UUuu11+Bm0SBR7dA+v1Fl7xnQTS/EYuHV2QIjv45LWN3a3xxWaIDiu2/VdJz4h2nivwAeU5N8oZBe5ClFu/IaQ=="
        )
    }
    
    @Test func testConfigureSheetPasswordAppliesLegacyAndModern() throws {
        var protection = SheetProtection()
        try CoreUtils.configureSheetPassword(
            &protection,
            plaintext: "789648",
            salt: Data("XLKitTestSalt!!".utf8)
        )
        #expect(protection.password == CoreUtils.excelLegacySheetPasswordHash(for: "789648"))
        #expect(protection.algorithmName == "SHA-512")
        #expect(protection.saltValue != nil)
        #expect(protection.hashValue != nil)
        #expect(protection.spinCount == 100_000)
    }
    
    @Test func testConfigureSheetPasswordRejectsEmptyPassword() {
        var protection = SheetProtection()
        #expect {
            try CoreUtils.configureSheetPassword(&protection, plaintext: "")
        } throws: { error in
            guard case .securityError(let message) = error as? XLKitError else {
                return false
            }
            return message.contains("must not be empty")
        }
    }
    
    @Test func testMinimalProtectionXML() {
        let protection = SheetProtection()
        #expect(XLSXEngine.sheetProtectionXML(protection) == "<sheetProtection sheet=\"1\"/>")
    }
    
    @Test func testProtectionWithLegacyPassword() {
        var protection = SheetProtection()
        protection.password = "CC3F"
        protection.selectLockedCells = true
        #expect(
            XLSXEngine.sheetProtectionXML(protection) == "<sheetProtection sheet=\"1\" password=\"CC3F\" selectLockedCells=\"1\"/>"
        )
    }
    
    @Test func testProtectionWithModernHash() {
        var protection = SheetProtection()
        protection.algorithmName = "SHA-512"
        protection.hashValue = "abc=="
        protection.saltValue = "def=="
        protection.spinCount = 100_000
        let xml = XLSXEngine.sheetProtectionXML(protection)
        #expect(xml.contains("algorithmName=\"SHA-512\""))
        #expect(xml.contains("hashValue=\"abc==\""))
        #expect(xml.contains("saltValue=\"def==\""))
        #expect(xml.contains("spinCount=\"100000\""))
    }
    
    @Test func testProtectionWithGranularPermissions() {
        var protection = SheetProtection()
        protection.formatCells = false
        protection.insertRows = false
        protection.sort = true
        let xml = XLSXEngine.sheetProtectionXML(protection)
        #expect(xml.contains("formatCells=\"0\""))
        #expect(xml.contains("insertRows=\"0\""))
        #expect(xml.contains("sort=\"1\""))
    }
    
    @Test func testProtectionOmittedWhenNotConfigured() {
        let sheet = Workbook().addSheet(name: "Sheet1")
        #expect(sheet.protection == nil)
    }
    
    @Test func testProtectionAllAttributesEmitted() {
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
            #expect(xml.contains("\(name)=\""), "Missing attribute \(name) in: \(xml)")
        }
    }
    
    @Test func testSavingWorkbookWithProtectedSheetSucceeds() throws {
        let workbook = Workbook()
        _ = workbook.addSheet(name: "Main")
        let tech = workbook.addSheet(name: "Strings")
        tech.protection = SheetProtection()
        let url = XLKitTestSupport.makeTempWorkbookURL(prefix: "protected_sheet")
        defer { XLKitTestSupport.cleanupTempFile(at: url) }
        try workbook.save(to: url)
        #expect(FileManager.default.fileExists(atPath: url.path))
    }
}
