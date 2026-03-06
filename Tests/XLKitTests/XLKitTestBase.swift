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
        // TimeZone(secondsFromGMT: 0) never returns nil, so force unwrap is safe
        calendar.timeZone = TimeZone(secondsFromGMT: 0)!
        
        guard let date = calendar.date(from: components) else {
            XCTFail("""
            Failed to create UTC date from components.
              Year:   \(components.year.map(String.init) ?? "nil")
              Month:  \(components.month.map(String.init) ?? "nil")
              Day:    \(components.day.map(String.init) ?? "nil")
              Hour:   \(components.hour.map(String.init) ?? "nil")
              Minute: \(components.minute.map(String.init) ?? "nil")
              Second: \(components.second.map(String.init) ?? "nil")
            Expected a valid Gregorian calendar date in UTC (e.g., year ≥ 1, month 1–12, day in valid range for the given month).
            """)
            // Return a deterministic fallback date instead of crashing the entire test suite.
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

    /// Helper to construct a standard test workbook with a single sheet and sample content.
    private func makeTestWorkbook() -> Workbook {
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Test")
        sheet.setCell("A1", value: .string("Test"))
        return workbook
    }

    /// Best-effort cleanup for temporary workbook files used in tests.
    /// Logs any failure via `XCTFail` so file system issues are visible in test output.
    private func cleanupTempFile(at url: URL) {
        do {
            try FileManager.default.removeItem(at: url)
        } catch {
            XCTFail("Failed to remove temporary workbook at \(url.path): \(error)")
        }
    }
    
    /// Helper to save a workbook to a temporary URL synchronously and ensure cleanup.
    func withSavedTempWorkbookSync(prefix: String,
                                   _ body: (_ workbook: Workbook, _ url: URL) throws -> Void) throws {
        let workbook = makeTestWorkbook()
        
        let tempURL = makeTempWorkbookURL(prefix: prefix)
        // Save the workbook to disk so callers can rely on a real file at tempURL.
        try workbook.save(to: tempURL)
        
        defer {
            cleanupTempFile(at: tempURL)
        }
        
        try body(workbook, tempURL)
    }
    
    /// Helper to save a workbook to a temporary URL asynchronously and ensure cleanup.
    func withSavedTempWorkbookAsync(prefix: String,
                                    _ body: @MainActor (_ workbook: Workbook, _ url: URL) async throws -> Void) async throws {
        let workbook = makeTestWorkbook()
        
        let tempURL = makeTempWorkbookURL(prefix: prefix)
        // Save the workbook to disk before invoking the async body.
        try await workbook.save(to: tempURL)
        
        defer {
            cleanupTempFile(at: tempURL)
        }
        
        try await body(workbook, tempURL)
    }
    
    /// Helper to create a bordered `CellFormat` with configurable top/bottom borders and optional color.
    private static func makeBorderedFormat(top: BorderStyle, bottom: BorderStyle, color: String? = nil) -> CellFormat {
        var format = CellFormat.bordered(style: .thin, color: color)
        format.borderTop = top
        format.borderBottom = bottom
        return format
    }

    /// Helper to create a `CellFormat` with thin borders on all sides.
    static func makeThinBorderFormat() -> CellFormat {
        makeBorderedFormat(top: .thin, bottom: .thin)
    }
    
    /// Helper to create a `CellFormat` with medium top/bottom borders and red border color.
    static func makeMediumRedBorderFormat() -> CellFormat {
        makeBorderedFormat(top: .medium, bottom: .medium, color: "#FF0000")
    }
    
    /// Helper to create a `CellFormat` with thick top/bottom borders and blue border color.
    static func makeThickBlueBorderFormat() -> CellFormat {
        makeBorderedFormat(top: .thick, bottom: .thick, color: "#0000FF")
    }
    
    /// Standard font size used in tests.
    static let standardFontSize: Double = 11
}
