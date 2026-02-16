//
//  XLKitTestBase.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

/// Base test class providing shared helper methods for all XLKit tests.
@MainActor
class XLKitTestBase: XCTestCase {
    
    /// Helper to construct a UTC date from calendar components, avoiding magic Unix timestamps.
    static func makeUTCDate(year: Int, month: Int, day: Int,
                           hour: Int = 0, minute: Int = 0, second: Int = 0) -> Date {
        var components = DateComponents()
        components.year = year
        components.month = month
        components.day = day
        components.hour = hour
        components.minute = minute
        components.second = second
        
        var calendar = Calendar(identifier: .gregorian)
        calendar.timeZone = TimeZone(secondsFromGMT: 0) ?? calendar.timeZone
        
        guard let date = calendar.date(from: components) else {
            XCTFail("Failed to create UTC date from components: \(components)")
            return Date(timeIntervalSince1970: 0)
        }
        
        return date
    }
    
    /// Fixed date used for deterministic date-related tests (2022-01-01 00:00:00 UTC).
    static let fixedTestDate = makeUTCDate(year: 2022, month: 1, day: 1)
    
    /// Epoch date (1970-01-01 00:00:00 UTC) used for simple date type tests.
    static let epochDate = makeUTCDate(year: 1970, month: 1, day: 1)
    
    /// Helper to generate a unique temporary URL for workbook files.
    /// - Parameter prefix: A prefix to distinguish different test cases.
    func makeTempWorkbookURL(prefix: String) -> URL {
        FileManager.default.temporaryDirectory.appendingPathComponent("\(prefix)-\(UUID().uuidString).xlsx")
    }
    
    /// Helper to save a workbook to a temporary URL synchronously and ensure cleanup.
    func withSavedTempWorkbookSync(prefix: String,
                                   _ body: (_ workbook: Workbook, _ url: URL) throws -> Void) throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        sheet.setCell("A1", value: .string("Test"))
        
        let tempURL = makeTempWorkbookURL(prefix: prefix)
        defer {
            // Best-effort cleanup; ignore errors to avoid hiding test failures.
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try body(workbook, tempURL)
    }
    
    /// Helper to save a workbook to a temporary URL asynchronously and ensure cleanup.
    func withSavedTempWorkbookAsync(prefix: String,
                                    _ body: (_ workbook: Workbook, _ url: URL) async throws -> Void) async throws {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        sheet.setCell("A1", value: .string("Test"))
        
        let tempURL = makeTempWorkbookURL(prefix: prefix)
        defer {
            // Best-effort cleanup; ignore errors to avoid hiding test failures.
            try? FileManager.default.removeItem(at: tempURL)
        }
        
        try await body(workbook, tempURL)
    }
    
    /// Helper to create a `CellFormat` with thin borders on all sides.
    static func makeThinBorderFormat() -> CellFormat {
        var format = CellFormat.bordered()
        format.borderTop = .thin
        format.borderBottom = .thin
        format.borderLeft = .thin
        format.borderRight = .thin
        return format
    }
    
    /// Helper to create a `CellFormat` with medium top/bottom borders and red border color.
    static func makeMediumRedBorderFormat() -> CellFormat {
        var format = CellFormat.bordered()
        format.borderTop = .medium
        format.borderBottom = .medium
        format.borderColor = "#FF0000"
        return format
    }
    
    /// Helper to create a `CellFormat` with thick top/bottom borders and blue border color.
    static func makeThickBlueBorderFormat() -> CellFormat {
        var format = CellFormat.bordered()
        format.borderTop = .thick
        format.borderBottom = .thick
        format.borderColor = "#0000FF"
        return format
    }
    
    /// Standard font size used in tests.
    static let standardFontSize: Double = 11
}
