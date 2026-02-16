//
//  FileOperationTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import XCTest
import XLKit

@MainActor
final class FileOperationTests: XLKitTestBase {
    
    func testWorkbookSave() throws {
        try withSavedTempWorkbookSync(prefix: "test") { workbook, url in
            // Test synchronous save
            try workbook.save(to: url)
            
            // Verify file exists
            XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
        }
    }
    
    func testWorkbookSaveAsync() async throws {
        try await withSavedTempWorkbookAsync(prefix: "test_async") { workbook, url in
            // Test asynchronous save
            try await workbook.save(to: url)
            
            // Verify file exists
            XCTAssertTrue(FileManager.default.fileExists(atPath: url.path))
        }
    }
}
