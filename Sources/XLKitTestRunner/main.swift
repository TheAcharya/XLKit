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

do {
    switch testType {
    case "no-embeds", "no-images":
        print("Executing: Generate Excel with No Image Embeds")
        print("[DEBUG] About to call generateExcelWithNoEmbeds()")
        try generateExcelWithNoEmbeds()
        print("[DEBUG] generateExcelWithNoEmbeds() completed")
        validateExcelFile("Test-Workflows/Embed-Test.xlsx")
    case "with-embeds", "with-images", "embed":
        print("Executing: Generate Excel with Image Embeds")
        print("[DEBUG] About to call generateExcelWithImageEmbeds()")
        try generateExcelWithImageEmbeds()
        print("[DEBUG] generateExcelWithImageEmbeds() completed")
        validateExcelFile("Test-Workflows/Embed-Test-Embed.xlsx")
    case "comprehensive", "demo":
        print("Executing: Comprehensive API Demonstration")
        print("[DEBUG] About to call demonstrateComprehensiveAPI()")
        try demonstrateComprehensiveAPI()
        print("[DEBUG] demonstrateComprehensiveAPI() completed")
        validateExcelFile("Test-Workflows/Comprehensive-Demo.xlsx")
    case "csv-import":
        print("[INFO] CSV import test not yet implemented")
    case "formatting":
        print("[INFO] Formatting test not yet implemented")
    case "help", "-h", "--help":
        printHelp()
    default:
        print("[ERROR] Unknown test type: \(testType)")
        printHelp()
    }
    print("[DEBUG] Execution completed successfully")
    exit(0)
} catch {
    print("[ERROR] Test failed: \(error)")
    print("[DEBUG] Execution failed with error")
    exit(1)
}

// MARK: - Help Function
func printHelp() {
    print("""
    
    Usage: swift run XLKitTestRunner [test-type]
    
    Available test types:
      no-embeds, no-images    - Generate Excel file from CSV without image embeds
      with-embeds, with-images, embed - Generate Excel file with image embeds
      comprehensive, demo     - Comprehensive API demonstration with all features
      csv-import              - Test CSV import functionality (future)
      formatting              - Test cell formatting features (future)
      help, -h, --help        - Show this help message
    
    Examples:
      swift run XLKitTestRunner no-embeds
      swift run XLKitTestRunner with-images
      swift run XLKitTestRunner comprehensive
      swift run XLKitTestRunner help
    
    
    """)
} 