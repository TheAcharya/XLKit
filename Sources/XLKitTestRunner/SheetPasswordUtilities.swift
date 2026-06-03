//
//  SheetPasswordUtilities.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit

/// Developer utility for worksheet protection password hashing (not used in CI workflows).
enum SheetPasswordUtilities {
    
    /// Prints legacy and modern hash values and a copy-paste Swift snippet for `SheetProtection`.
    /// - Parameters:
    ///   - password: Plaintext worksheet password.
    ///   - useComprehensiveDemoSalts: When true and password is `1234`, also prints hashes for both comprehensive-demo sheets.
    static func printPasswordHashes(for password: String, useComprehensiveDemoSalts: Bool = false) throws {
        print("XLKit Sheet Protection Password Helper")
        print("====================================")
        print("")
        print("Plaintext password: \(password)")
        print("")
        
        let legacy = CoreUtils.excelLegacySheetPasswordHash(for: password)
        print("Legacy (SheetProtection.password):")
        print("  \(legacy)")
        print("")
        
        let modern = try CoreUtils.excelModernSheetPasswordHash(for: password)
        printModernFields(modern, title: "Modern SHA-512 (random salt):")
        printSwiftSnippet(password: password, legacy: legacy, modern: modern, legacyInSnippet: true)
        
        if useComprehensiveDemoSalts && password == CoreUtils.comprehensiveDemoSheetPassword {
            print("")
            print("Comprehensive demo salts (reproducible Comprehensive-Demo.xlsx):")
            print("")
            
            let passwordSheetModern = try CoreUtils.excelModernSheetPasswordHash(
                for: password,
                salt: CoreUtils.comprehensiveDemoPasswordSheetSalt
            )
            printModernFields(passwordSheetModern, title: "  Protected (Password) sheet:")
            
            let modernOnly = try CoreUtils.excelModernSheetPasswordHash(
                for: password,
                salt: CoreUtils.comprehensiveDemoModernSheetSalt
            )
            printModernFields(modernOnly, title: "  Protected (Modern Hash) sheet — legacy omitted:")
            print("")
            print("  Use CoreUtils.comprehensiveDemoPasswordSheetSalt / comprehensiveDemoModernSheetSalt in code.")
        }
        
        print("")
        print("In Excel: type the plaintext password above to unprotect the sheet.")
    }
    
    private static func printModernFields(_ modern: CoreUtils.ExcelModernSheetPasswordHash, title: String) {
        print(title)
        print("  algorithmName: \(modern.algorithmName)")
        print("  saltValue:     \(modern.saltValue)")
        print("  hashValue:     \(modern.hashValue)")
        print("  spinCount:     \(modern.spinCount)")
        print("")
    }
    
    private static func printSwiftSnippet(
        password: String,
        legacy: String,
        modern: CoreUtils.ExcelModernSheetPasswordHash,
        legacyInSnippet: Bool
    ) {
        print("Swift snippet:")
        print("--------------")
        var lines = [
            "var protection = SheetProtection()",
        ]
        if legacyInSnippet {
            lines.append("protection.password = CoreUtils.excelLegacySheetPasswordHash(for: \"\(password)\") // \(legacy)")
        }
        lines.append("protection.algorithmName = \"\(modern.algorithmName)\"")
        lines.append("protection.saltValue = \"\(modern.saltValue)\"")
        lines.append("protection.hashValue = \"\(modern.hashValue)\"")
        lines.append("protection.spinCount = \(modern.spinCount)")
        lines.append("sheet.protection = protection")
        print(lines.joined(separator: "\n"))
        print("")
        print("Or:")
        print("try CoreUtils.configureSheetPassword(&protection, plaintext: \"\(password)\")")
    }
}
