//
//  main.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation

// MARK: - Command Line Interface

print("XLKit Test Runner")
print("==================")

// Get command line arguments
let arguments = CommandLine.arguments
let testType = arguments.count > 1 ? arguments[1] : "no-embeds"

print("Running test type: \(testType)")

// MARK: - Test Runner

switch testType {
case "no-embeds", "no-images":
    print("Executing: Generate Excel with No Image Embeds")
    generateExcelWithNoEmbeds()
    
case "with-embeds", "with-images":
    print("Executing: Generate Excel with Image Embeds")
    // generateExcelWithEmbeds() // Future implementation
    
case "csv-import":
    print("Executing: CSV Import Test")
    // csvImportTest() // Future implementation
    
case "formatting":
    print("Executing: Cell Formatting Test")
    // cellFormattingTest() // Future implementation
    
case "help", "-h", "--help":
    printHelp()
    
default:
    print("Unknown test type: \(testType)")
    printHelp()
    exit(1)
}

// MARK: - Help Function

func printHelp() {
    print("""
    
    Usage: swift run XLKitTestRunner [test-type]
    
    Available test types:
      no-embeds, no-images    - Generate Excel file from CSV without image embeds
      with-embeds, with-images - Generate Excel file with image embeds (future)
      csv-import              - Test CSV import functionality (future)
      formatting              - Test cell formatting features (future)
      help, -h, --help        - Show this help message
    
    Examples:
      swift run XLKitTestRunner no-embeds
      swift run XLKitTestRunner with-images
      swift run XLKitTestRunner help
    
    """)
} 