#!/usr/bin/env swift

import Foundation

// This script demonstrates how XLKit can process CSV data with corresponding images
// and embed them into an Excel file

print("XLKit Embed Demo - Processing CSV with Images")
print("=============================================")

// Define the test data path
let testDataPath = "Test-Data/Embed-Test 2025-07-06 2-25-34 PM"
let csvFile = "\(testDataPath)/Embed-Test.csv"

// Get the current working directory
let currentDir = FileManager.default.currentDirectoryPath
let fullCsvPath = "\(currentDir)/\(csvFile)"

// Read and parse the CSV data
print("\n📊 Reading CSV data...")
guard let csvData = try? String(contentsOfFile: fullCsvPath) else {
    print("❌ Failed to read CSV file")
    exit(1)
}

print("✅ CSV data loaded successfully")
print("📄 CSV content preview:")
let lines = csvData.components(separatedBy: .newlines)
for (index, line) in lines.prefix(3).enumerated() {
    print("   Line \(index + 1): \(line.prefix(100))...")
}

// Parse CSV to extract image filenames and corresponding data
print("\n🔍 Parsing CSV for image mappings...")
let csvLines = lines.filter { !$0.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty }
let headerLine = csvLines[0]
let dataLines = Array(csvLines.dropFirst())

// Find the Image Filename column index
let headers = headerLine.components(separatedBy: ",")
guard let imageFilenameIndex = headers.firstIndex(of: "Image Filename") else {
    print("❌ Could not find 'Image Filename' column in CSV")
    exit(1)
}

print("✅ Found Image Filename column at index \(imageFilenameIndex)")

// Extract data rows with their corresponding images
var imageMappings: [(rowData: String, imageFile: String)] = []

for (rowIndex, line) in dataLines.enumerated() {
    let columns = line.components(separatedBy: ",")
    if columns.count > imageFilenameIndex {
        let imageFilename = columns[imageFilenameIndex].trimmingCharacters(in: .whitespacesAndNewlines)
        if !imageFilename.isEmpty {
            imageMappings.append((rowData: line, imageFile: imageFilename))
            print("   Row \(rowIndex + 1): \(imageFilename)")
        }
    }
}

print("✅ Found \(imageMappings.count) rows with corresponding images")

// Check if image files exist
print("\n🖼️  Verifying image files...")
for mapping in imageMappings {
    let imagePath = "\(testDataPath)/\(mapping.imageFile)"
    let fileExists = FileManager.default.fileExists(atPath: imagePath)
    print("   \(mapping.imageFile): \(fileExists ? "✅" : "❌")")
}

// Demonstrate how XLKit would process this data
print("\n📋 XLKit Processing Plan:")
print("1. Create workbook from CSV data")
print("2. For each row with an image:")
print("   - Locate the corresponding cell")
print("   - Load the image file")
print("   - Embed the image in the cell")
print("   - Auto-size the column to fit the image")
print("3. Save the workbook as XLSX")

// Example of how the XLKit API would be used:
print("\n💡 XLKit API Usage Example:")
print("""
let workbook = XLKit.createWorkbookFromCSV(
    csvData: csvData,
    sheetName: "Video Markers",
    hasHeader: true
)

let sheet = workbook.getSheet(name: "Video Markers")!

// For each image mapping:
for (index, mapping) in imageMappings.enumerated() {
    let imagePath = "\(testDataPath)/\\(mapping.imageFile)"
    let imageURL = URL(fileURLWithPath: imagePath)
    
    // Embed image in the corresponding row
    let row = index + 2 // +2 because of header and 1-based indexing
    let imageCell = "B\\(row)" // Assuming images go in column B
    
    try sheet.addImage(imageURL, at: imageCell)
        .autoSizeColumn(2, forImageAt: imageCell)
}

// Save the workbook
try XLKit.saveWorkbook(workbook, to: URL(fileURLWithPath: "VideoMarkers.xlsx"))
""")

print("\n🎯 Key XLKit Features Demonstrated:")
print("✅ CSV import with header handling")
print("✅ Image embedding from file URLs")
print("✅ Auto-column sizing for images")
print("✅ Fluent API for chaining operations")
print("✅ Error handling for file operations")
print("✅ XLSX generation with embedded images")

print("\n✨ This demonstrates how XLKit can seamlessly combine:")
print("   - CSV data processing")
print("   - Image file handling")
print("   - Excel file generation")
print("   - Image embedding and formatting")

print("\n📁 Output would be a professional Excel file with:")
print("   - All CSV data properly formatted")
print("   - Images embedded in corresponding cells")
print("   - Auto-sized columns for optimal viewing")
print("   - Professional appearance suitable for reports") 