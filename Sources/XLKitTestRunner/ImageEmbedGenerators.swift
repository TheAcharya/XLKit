//
//  ImageEmbedGenerators.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit
import CoreXLSX

// MARK: - ImageEmbedGenerators

/// ImageEmbedGenerators
///
/// Provides test routines for embedding images into Excel files using XLKit.
/// This module demonstrates pixel-perfect image embedding, automatic sizing,
/// and Excel compliance validation using CoreXLSX. Used by the test runner
/// to validate image embedding features and aspect ratio preservation.
///
/// ## Features Demonstrated
/// - CSV-driven Excel file generation
/// - Automatic column width calculation
/// - Bold header formatting
/// - Image embedding with perfect aspect ratio preservation
/// - Automatic cell sizing for images
/// - Output to `Test-Workflows/Embed-Test-Embed.xlsx`
/// - CoreXLSX validation for Excel compliance
struct ImageEmbedGenerators {
    /// Generates an Excel file with images embedded in a dedicated column.
    ///
    /// - Reads data from a CSV file in `Test-Data/Embed-Test/Embed-Test.csv`
    /// - Embeds all images found in `Test-Data/Embed-Test/` (PNG, JPG, JPEG, GIF)
    /// - Applies bold formatting to headers and auto-sizes columns
    /// - Embeds each image in a new row, preserving aspect ratio and sizing
    /// - Saves the result to `Test-Workflows/Embed-Test-Embed.xlsx`
    /// - Validates the output file using CoreXLSX
    ///
    /// Throws: `XLKitError` if file operations or embedding fail
    @MainActor
    static func generateExcelWithImageEmbeds() async throws {
        print("[INFO] Starting image embedding test...")
        
        // MARK: - Configuration
        let csvFilePath = "Test-Data/Embed-Test/Embed-Test.csv"
        let outputExcelFile = "Test-Workflows/Embed-Test-Embed.xlsx"
        
        print("[INFO] Using CSV file: \(csvFilePath)")
        
        // Ensure CSV file exists
        guard FileManager.default.fileExists(atPath: csvFilePath) else {
            print("[ERROR] CSV file not found: \(csvFilePath)")
            throw XLKitError.fileWriteError("CSV file not found: \(csvFilePath)")
        }
        
        // MARK: - CSV Parsing
        /// Parses a CSV string into rows and headers.
        /// - Parameter csv: The CSV file contents as a string.
        /// - Returns: Tuple of (rows, headers)
        func parseCSV(_ csv: String) -> (rows: [[String]], headers: [String]) {
            let lines = csv.components(separatedBy: .newlines).filter { !$0.isEmpty }
            guard let headerLine = lines.first else { return ([], []) }
            let headers = headerLine.components(separatedBy: ",")
            let rows = lines.dropFirst().map { $0.components(separatedBy: ",") }
            return (rows, headers)
        }
        
        guard let csvData = try? String(contentsOfFile: csvFilePath, encoding: .utf8) else {
            print("[ERROR] Failed to read CSV file: \(csvFilePath)")
            throw XLKitError.fileWriteError("Failed to read CSV file: \(csvFilePath)")
        }
        
        let (rows, headers) = parseCSV(csvData)
        print("[INFO] Parsed CSV with \(headers.count) columns and \(rows.count) data rows")
        
        // MARK: - Create Excel File with XLKit
        print("[INFO] Creating Excel workbook...")
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Embed Test")
        
        // MARK: - Calculate Column Widths
        print("[INFO] Calculating optimal column widths...")
        var columnWidths: [Int: Double] = [:]
        
        // Calculate width for each column based on longest value (header or data)
        for col in 0..<headers.count {
            let column = col + 1
            var maxWidth = calculateTextWidth(headers[col])
            for row in rows {
                if col < row.count {
                    let cellWidth = calculateTextWidth(row[col])
                    maxWidth = max(maxWidth, cellWidth)
                }
            }
            // Add padding, clamp to min/max
            let adjustedWidth = min(max(maxWidth + 4.0, 8.0), 50.0)
            columnWidths[column] = adjustedWidth
            print("[INFO] Column \(column) (\(headers[col])): width = \(adjustedWidth)")
        }
        
        // MARK: - Write Headers with Bold Formatting
        print("[INFO] Writing headers with bold formatting...")
        let headerFormat = CellFormat.header(fontSize: 12.0, backgroundColor: "#E6E6E6")
        for (col, header) in headers.enumerated() {
            let column = col + 1
            sheet.setCell(row: 1, column: column, cell: Cell.string(header, format: headerFormat))
        }
        
        // MARK: - Write Data Rows
        print("[INFO] Writing data rows...")
        for (rowIdx, row) in rows.enumerated() {
            let excelRow = rowIdx + 2
            for (colIdx, value) in row.enumerated() {
                let column = colIdx + 1
                sheet.setCell(row: excelRow, column: column, value: .string(value))
            }
        }
        
        // MARK: - Apply Column Widths
        print("[INFO] Applying column widths...")
        for (column, width) in columnWidths {
            sheet.setColumnWidth(column, width: width)
        }
        
        // MARK: - Image Embedding
        print("[INFO] Starting image embedding process...")
        let imageDirectory = "Test-Data/Embed-Test"
        let imageFiles = try getImageFiles(from: imageDirectory)
        print("[INFO] Found \(imageFiles.count) image files for embedding")
        
        // Add header for image column with consistent formatting
        let imageEmbedColumn = headers.count + 1
        sheet.setCell(row: 1, column: imageEmbedColumn, cell: Cell.string("Image", format: headerFormat))

        // Embed each image in a new row, preserving aspect ratio
        for (index, imageFile) in imageFiles.enumerated() {
            let row = index + 2 // Start from row 2 (after header)
            let coordinate = CellCoordinate(row: row, column: imageEmbedColumn).excelAddress
            print("[INFO] Embedding image \(index + 1)/\(imageFiles.count): \(imageFile.lastPathComponent) at \(coordinate)")
            do {
                let imageData = try Data(contentsOf: imageFile)
                // Use the XLKit API for image embedding with automatic sizing
                let success = try await sheet.embedImageAutoSized(
                    imageData,
                    at: coordinate,
                    of: workbook,
                    scale: 0.5 // 50% scaling for compact images
                )
                if success {
                    print("[INFO] ✓ Successfully embedded image at \(coordinate)")
                } else {
                    print("[WARNING] Failed to embed image at \(coordinate)")
                }
            } catch {
                print("[ERROR] Failed to embed image \(imageFile.lastPathComponent): \(error)")
                // Continue with other images
            }
        }
        
        // MARK: - Save and Validate
        print("[INFO] Saving Excel file with embedded images...")
        let outputURL = URL(fileURLWithPath: outputExcelFile)
        let outputDir = outputURL.deletingLastPathComponent()
        if !FileManager.default.fileExists(atPath: outputDir.path) {
            print("[INFO] Creating output directory: \(outputDir.path)")
            do {
                try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)
            } catch {
                print("[ERROR] Failed to create output directory: \(error)")
                throw XLKitError.fileWriteError("Failed to create output directory: \(error)")
            }
        }
        do {
            try await workbook.save(to: outputURL)
            print("[SUCCESS] Excel file with embedded images created: \(outputExcelFile)")
            print("[INFO] Features applied:")
            print("  - Bold headers with gray background")
            print("  - Automatic column width adjustment")
            print("  - \(headers.count) columns optimized")
            print("  - \(imageFiles.count) images embedded with perfect aspect ratio preservation")
            print("  - Automatic cell sizing for images")
        } catch {
            print("[ERROR] Failed to save Excel file: \(error)")
            throw error
        }
        
        // MARK: - CoreXLSX Validation
        print("[INFO] Validating with CoreXLSX...")
        do {
            let xlsx = XLSXFile(filepath: outputExcelFile)
            let worksheet = try xlsx?.parseWorksheet(at: "xl/worksheets/sheet1.xml")
            let sharedStrings = try xlsx?.parseSharedStrings()
            print("[INFO] ✓ CoreXLSX validation successful")
            print("[INFO] ✓ Worksheet contains \(worksheet?.data?.rows.count ?? 0) rows")
            print("[INFO] ✓ Shared strings count: \(sharedStrings?.items.count ?? 0)")
        } catch {
            print("[ERROR] CoreXLSX validation failed: \(error)")
            throw error
        }
        print("[SUCCESS] Image embedding test completed successfully!")
    }

    // MARK: - Helper Functions

    /// Calculates an approximate text width for column sizing.
    /// - Parameter text: The text to measure.
    /// - Returns: Approximate width in pixels (8 pixels per character).
    private static func calculateTextWidth(_ text: String) -> Double {
        return Double(text.count) * 8.0
    }

    /// Returns a sorted list of image file URLs from a directory.
    /// - Parameter directory: Directory path to search for images.
    /// - Returns: Array of image file URLs (PNG, JPG, JPEG, GIF).
    /// - Throws: `XLKitError` if the directory does not exist or cannot be read.
    private static func getImageFiles(from directory: String) throws -> [URL] {
        let directoryURL = URL(fileURLWithPath: directory)
        guard FileManager.default.fileExists(atPath: directory) else {
            throw XLKitError.fileWriteError("Directory not found: \(directory)")
        }
        let contents = try FileManager.default.contentsOfDirectory(at: directoryURL, includingPropertiesForKeys: nil)
        let imageExtensions = ["png", "jpg", "jpeg", "gif"]
        let imageFiles = contents.filter { url in
            let fileExtension = url.pathExtension.lowercased()
            return imageExtensions.contains(fileExtension)
        }
        return imageFiles.sorted { $0.lastPathComponent < $1.lastPathComponent }
    }
} 