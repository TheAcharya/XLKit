//
//  main.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation

// MARK: - Command Line Interface
print("XLKit Test Runner")
print("==================")

let args = CommandLine.arguments

if args.count < 2 {
    printHelp()
    exit(0)
}

let testType = args[1].lowercased()
print("Running test type: \(testType)")

print("[DEBUG] Starting execution...")

var exitCode: Int32 = 0

let mainTask = Task {
    do {
        switch testType {
        case "no-embeds", "no-images":
            print("Executing: Generate Excel with No Image Embeds")
            print("[DEBUG] About to call generateExcelWithNoEmbeds()")
            try ExcelGenerators.generateExcelWithNoEmbeds()
            print("[DEBUG] generateExcelWithNoEmbeds() completed")
            ExcelGenerators.validateExcelFile("Test-Workflows/Embed-Test.xlsx")
        case "with-embeds", "with-images", "embed":
            print("Executing: Generate Excel with Image Embeds")
            print("[DEBUG] About to call generateExcelWithImageEmbeds()")
            try await ImageEmbedGenerators.generateExcelWithImageEmbeds()
            print("[DEBUG] generateExcelWithImageEmbeds() completed")
        case "comprehensive", "demo":
            print("Executing: Comprehensive API Demonstration")
            print("[DEBUG] About to call demonstrateComprehensiveAPI()")
            try await ExcelGenerators.demonstrateComprehensiveAPI()
            print("[DEBUG] demonstrateComprehensiveAPI() completed")
        case "security-demo", "security":
            print("Executing: File Path Security Restrictions Demo")
            print("[DEBUG] About to call demonstrateFilePathRestrictions()")
            ExcelGenerators.demonstrateFilePathRestrictions()
            print("[DEBUG] demonstrateFilePathRestrictions() completed")
        case "ios-test", "ios":
            print("Executing: iOS Compatibility Test")
            print("[DEBUG] About to call testIOSCompatibility()")
            try await ExcelGenerators.testIOSCompatibility()
            print("[DEBUG] testIOSCompatibility() completed")
        case "help", "-h", "--help":
            printHelp()
        default:
            print("[ERROR] Unknown test type: \(testType)")
            printHelp()
        }
        print("[DEBUG] Execution completed successfully")
        exitCode = 0
    } catch {
        print("[ERROR] Test failed: \(error)")
        print("[DEBUG] Execution failed with error")
        exitCode = 1
    }
}

// Wait for the async main task to finish
import _Concurrency
_ = await mainTask.value
exit(exitCode)

// MARK: - Help Function
func printHelp() {
    print("""
    
    Usage: swift run XLKitTestRunner [test-type]
    
    Available test types:
      no-embeds, no-images    - Generate Excel file from CSV without image embeds
      with-embeds, with-images, embed - Generate Excel file with image embeds
      comprehensive, demo     - Comprehensive API demonstration
      security-demo, security - Demonstrate file path security restrictions
      ios-test, ios           - Test iOS file system compatibility
      help, -h, --help        - Show this help message
    
    Examples:
      swift run XLKitTestRunner no-embeds
      swift run XLKitTestRunner security-demo
      swift run XLKitTestRunner ios-test
      swift run XLKitTestRunner help
    
    
    """)
} 