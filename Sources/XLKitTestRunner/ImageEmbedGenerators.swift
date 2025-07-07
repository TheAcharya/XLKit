//
//  ImageEmbedGenerators.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit
import CoreXLSX

// MARK: - Image Embedding Generation Functions

func generateExcelWithImageEmbeds() {
    print("[INFO] Starting Excel generation with image embeds...")
    
    // Use the same CSV data as base
    let csvPath = "Test-Data/Embed-Test/Embed-Test.csv"
    print("[INFO] Using CSV file: \(csvPath)")
    
    guard let csvData = try? String(contentsOfFile: csvPath, encoding: .utf8) else {
        print("[ERROR] Failed to read CSV file")
        exit(1)
    }
    
    // Parse CSV
    let lines = csvData.components(separatedBy: .newlines).filter { !$0.isEmpty }
    guard lines.count > 1 else {
        print("[ERROR] CSV file is empty or has no data rows")
        exit(1)
    }
    
    let headers = lines[0].components(separatedBy: ",").map { $0.trimmingCharacters(in: .whitespacesAndNewlines) }
    let dataRows = Array(lines.dropFirst())
    
    print("[INFO] Parsed CSV with \(headers.count) columns and \(dataRows.count) data rows")
    
    // Find the Image Filename column index
    guard let imageFilenameColumnIndex = headers.firstIndex(of: "Image Filename") else {
        print("[ERROR] 'Image Filename' column not found in CSV")
        exit(1)
    }
    
    print("[INFO] Found 'Image Filename' column at index \(imageFilenameColumnIndex)")
    
    // Create workbook using the improved API
    print("[INFO] Creating Excel workbook...")
    let workbook = XLKit.createWorkbook()
    let sheet = workbook.addSheet(name: "Embed Test")
    
    // Create new headers with "Image" column after "Image Filename"
    var newHeaders = headers
    newHeaders.insert("Image", at: imageFilenameColumnIndex + 1)
    
    print("[INFO] Created new headers with 'Image' column: \(newHeaders)")
    
    // Calculate optimal column widths
    print("[INFO] Calculating optimal column widths...")
    var columnWidths: [Double] = []
    
    for (index, header) in newHeaders.enumerated() {
        var maxWidth = calculateTextWidth(header) + 2.0 // Add padding
        
        // Check data rows for this column (adjust for new column structure)
        for row in dataRows {
            let columns = row.components(separatedBy: ",")
            if index < columns.count {
                let cellValue = columns[index].trimmingCharacters(in: .whitespacesAndNewlines)
                let cellWidth = calculateTextWidth(cellValue)
                maxWidth = max(maxWidth, cellWidth + 2.0)
            }
        }
        
        // For Image column, let the image embedding method set the optimal width
        if header == "Image" {
            // Don't set a fixed width - let embedImageAutoSized calculate it
            maxWidth = max(maxWidth, 20.0) // Minimum width for the column header
        }
        
        columnWidths.append(maxWidth)
        print("[INFO] Column \(index + 1) (\(header)): width = \(maxWidth)")
    }
    
    // Write headers with bold formatting using the improved API
    print("[INFO] Writing headers with bold formatting...")
    for (index, header) in newHeaders.enumerated() {
        let coordinate = "\(CoreUtils.columnLetter(from: index + 1))1"
        let headerFormat = CellFormat.header(fontSize: 12.0, backgroundColor: "#E6E6E6")
        sheet.setCell(coordinate, string: header, format: headerFormat)
    }
    
    // Write data rows and embed images using the improved API
    print("[INFO] Writing data rows and embedding images...")
    let imageDirectory = "Test-Data/Embed-Test/"
    
    for (rowIndex, row) in dataRows.enumerated() {
        let excelRow = rowIndex + 2 // Start from row 2 (row 1 is headers)
        let columns = row.components(separatedBy: ",")
        
        // Write all original columns
        for (colIndex, cellValue) in columns.enumerated() {
            let coordinate = "\(CoreUtils.columnLetter(from: colIndex + 1))\(excelRow)"
            let trimmedValue = cellValue.trimmingCharacters(in: .whitespacesAndNewlines)
            sheet.setCell(coordinate, value: .string(trimmedValue))
        }
        
        // Handle the new "Image" column
        let imageColumnIndex = imageFilenameColumnIndex + 1
        let imageColumnCoordinate = "\(CoreUtils.columnLetter(from: imageColumnIndex + 1))\(excelRow)"
        
        // Get the image filename from the Image Filename column
        if imageFilenameColumnIndex < columns.count {
            let imageFilename = columns[imageFilenameColumnIndex].trimmingCharacters(in: .whitespacesAndNewlines)
            
            if !imageFilename.isEmpty {
                let imagePath = imageDirectory + imageFilename
                print("[INFO] Processing row \(excelRow): \(imageFilename)")
                do {
                    let imageData = try Data(contentsOf: URL(fileURLWithPath: imagePath))
                    
                    // Debug: Check image format and size
                    let detectedFormat = ImageUtils.detectImageFormat(from: imageData)
                    let originalSize = ImageUtils.getImageSize(from: imageData, format: detectedFormat ?? .png)
                    let sizeString = originalSize != nil ? "\(originalSize!.width)x\(originalSize!.height)" : "unknown"
                    print("[DEBUG] Image: \(imageFilename), Format: \(detectedFormat?.rawValue ?? "unknown"), Size: \(sizeString)")
                    
                    // Use the improved embedImageAutoSized method for dynamic sizing
                    // Use larger bounds to compensate for the 0.5 scale factor
                    let success = sheet.embedImageAutoSized(
                        imageData,
                        at: imageColumnCoordinate,
                        of: workbook,
                        format: detectedFormat,
                        maxCellWidth: 1200, // Double the default to compensate for 0.5 scale
                        maxCellHeight: 800   // Double the default to compensate for 0.5 scale
                    )
                    
                    if success {
                        // Don't set any cell value - just embed the image
                        print("[INFO] ✓ Embedded \(imageFilename) at \(imageColumnCoordinate) with auto-sizing")
                    } else {
                        print("[WARNING] Failed to embed \(imageFilename)")
                        sheet.setCell(imageColumnCoordinate, value: .string("Embed failed: \(imageFilename)"))
                    }
                } catch {
                    print("[WARNING] Failed to process image \(imageFilename): \(error)")
                    sheet.setCell(imageColumnCoordinate, value: .string("Error: \(error.localizedDescription)"))
                }
            } else {
                sheet.setCell(imageColumnCoordinate, value: .string(""))
            }
        }
    }
    
    // Apply column widths (but don't override image column widths set by embedImageAutoSized)
    print("[INFO] Applying column widths...")
    for (index, width) in columnWidths.enumerated() {
        let columnNumber = index + 1
        // Skip the Image column as it's already set by embedImageAutoSized
        if index != imageFilenameColumnIndex + 1 {
            sheet.setColumnWidth(columnNumber, width: width)
        }
    }
    
    // Demonstrate workbook image management
    print("[INFO] Workbook contains \(workbook.imageCount) images")
    let pngImages = workbook.getImages(withFormat: .png)
    print("[INFO] Found \(pngImages.count) PNG images in workbook")
    
    // Save Excel file using the improved API
    print("[INFO] Saving Excel file...")
    let outputPath = "Test-Workflows/Embed-Test-Embed.xlsx"
    
    // Ensure output directory exists
    let outputURL = URL(fileURLWithPath: outputPath)
    let outputDir = outputURL.deletingLastPathComponent()
    if !FileManager.default.fileExists(atPath: outputDir.path) {
        print("[INFO] Creating output directory: \(outputDir.path)")
        do {
            try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
        } catch {
            print("[ERROR] Failed to create output directory: \(error)")
            exit(1)
        }
    }
    
    do {
        try XLKit.saveWorkbook(workbook, to: URL(fileURLWithPath: outputPath))
        print("[SUCCESS] Excel file created: \(outputPath)")
        print("[INFO] Features applied:")
        print("  - Bold headers with gray background")
        print("  - Automatic column width adjustment")
        print("  - \(newHeaders.count) columns optimized")
        print("  - \(workbook.imageCount) images embedded using improved API")
        print("  - Images positioned in cells with proper sizing")
        
    } catch {
        print("[ERROR] Failed to save Excel file: \(error)")
        exit(1)
    }
}

// MARK: - Helper Functions

private func calculateTextWidth(_ text: String) -> Double {
    // Simple approximation: each character is roughly 1.2 units wide
    return Double(text.count) * 1.2
} 