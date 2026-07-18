//
//  FileOperationTests.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import Testing
import XLKit

@Suite
@MainActor
struct FileOperationTests {
    
    @Test func testWorkbookSave() throws {
        try XLKitTestSupport.withSavedTempWorkbookSync(prefix: "test") { workbook, url in
            // Test synchronous save
            try workbook.save(to: url)
            
            // Verify file exists
            #expect(FileManager.default.fileExists(atPath: url.path))
        }
    }
    
    @Test func testWorkbookSaveAsync() async throws {
        try await XLKitTestSupport.withSavedTempWorkbookAsync(prefix: "test_async") { workbook, url in
            // Test asynchronous save
            try await workbook.save(to: url)
            
            // Verify file exists
            #expect(FileManager.default.fileExists(atPath: url.path))
        }
    }
}
