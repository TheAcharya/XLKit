//
//  ComprehensiveDemoProtection.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2026 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation

/// Demo password and salts for the `comprehensive` CLI generator (`Comprehensive-Demo.xlsx`).
enum ComprehensiveDemoProtection {
    static let password = "1234"
    static let passwordSheetSalt = Data("XLKitPassDemo1!!".utf8)
    static let modernSheetSalt = Data("XLKitModHashDemo!".utf8)
}
