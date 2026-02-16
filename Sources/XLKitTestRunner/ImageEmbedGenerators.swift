//
//  ImageEmbedGenerators.swift
//  XLKitTestRunner • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import XLKit
import CoreXLSX

/// ImageEmbedGenerators
///
/// Provides test routines for embedding images into Excel files using XLKit.
/// Demonstrates pixel-perfect image embedding, automatic sizing, and Excel
/// compliance validation using CoreXLSX.
///
/// Features:
/// - CSV-driven Excel file generation
/// - Automatic column width calculation
/// - Bold header formatting
/// - Image embedding with perfect aspect ratio preservation
/// - Automatic cell sizing for images
/// - Output to `Test-Workflows/Embed-Test-Embed.xlsx`
/// - CoreXLSX validation for Excel compliance
struct ImageEmbedGenerators {
    /// Generates Excel file with images embedded in a dedicated column
    ///
    /// - Reads data from CSV file in `Test-Data/Embed-Test/Embed-Test.csv`
    /// - Embeds all images found in `Test-Data/Embed-Test/` (PNG, JPG, JPEG, GIF)
    /// - Applies bold formatting to headers and auto-sizes columns
    /// - Embeds each image in a new row, preserving aspect ratio and sizing
    /// - Saves to `Test-Workflows/Embed-Test-Embed.xlsx`
    /// - Validates output file using CoreXLSX
    @MainActor
    static func generateExcelWithImageEmbeds() async throws {
        print("[INFO] Starting image embedding test...")
        
        // MARK: - Configuration
        let csvFilePath = "Test-Data/Embed-Test/Embed-Test.csv"
        let outputExcelFile = "Test-Workflows/Embed-Test-Embed.xlsx"
        
        print("[INFO] Using CSV file: \(csvFilePath)")
        
        guard FileManager.default.fileExists(atPath: csvFilePath) else {
            print("[ERROR] CSV file not found: \(csvFilePath)")
            throw XLKitError.fileWriteError("CSV file not found: \(csvFilePath)")
        }

        guard let csvData = try? String(contentsOfFile: csvFilePath, encoding: .utf8) else {
            print("[ERROR] Failed to read CSV file: \(csvFilePath)")
            throw XLKitError.fileWriteError("Failed to read CSV file: \(csvFilePath)")
        }
        
        // MARK: - Create Excel File
        print("[INFO] Creating Excel workbook...")
        let workbook = Workbook()
        let sheet = workbook.addSheet(name: "Embed Test")

        // Import CSV using XLKit's public API (powered by swift-textfile-tools for robust parsing).
        // Use hasHeader: false so the header row is included in the sheet and can be formatted.
        workbook.importCSV(csvData, into: sheet, hasHeader: false)

        let usedCoordinates = sheet.getUsedCells().compactMap { CellCoordinate(excelAddress: $0) }
        let maxRow = usedCoordinates.map(\.row).max() ?? 0
        let maxColumn = usedCoordinates.map(\.column).max() ?? 0

        let headers: [String] = (1...maxColumn).map { column in
            let coordinate = CellCoordinate(row: 1, column: column).excelAddress
            return sheet.getCell(coordinate)?.stringValue ?? ""
        }

        let dataRowCount = max(0, maxRow - 1)
        print("[INFO] Imported CSV with \(headers.count) columns and \(dataRowCount) data rows")
        
        // MARK: - Calculate Column Widths
        print("[INFO] Calculating optimal column widths...")
        var columnWidths: [Int: Double] = [:]
        
        for col in 1...maxColumn {
            var maxWidth = calculateTextWidth(headers[col - 1])
            for row in 2...maxRow {
                let coordinate = CellCoordinate(row: row, column: col).excelAddress
                let text = sheet.getCell(coordinate)?.stringValue ?? ""
                maxWidth = max(maxWidth, calculateTextWidth(text))
            }
            let adjustedWidth = min(max(maxWidth + 4.0, 8.0), 50.0)
            columnWidths[col] = adjustedWidth
            print("[INFO] Column \(col) (\(headers[col - 1])): width = \(adjustedWidth)")
        }
        
        // MARK: - Write Headers
        print("[INFO] Writing headers with bold formatting...")
        let headerFormat = CellFormat.header(fontSize: 12.0, backgroundColor: "#E6E6E6")
        for (col, header) in headers.enumerated() {
            let column = col + 1
            sheet.setCell(row: 1, column: column, cell: Cell.string(header, format: headerFormat))
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
        
        // Add header for image column
        let imageEmbedColumn = headers.count + 1
        sheet.setCell(row: 1, column: imageEmbedColumn, cell: Cell.string("Image", format: headerFormat))

        // Embed each image in a new row
        for (index, imageFile) in imageFiles.enumerated() {
            let row = index + 2
            let coordinate = CellCoordinate(row: row, column: imageEmbedColumn).excelAddress
            print("[INFO] Embedding image \(index + 1)/\(imageFiles.count): \(imageFile.lastPathComponent) at \(coordinate)")
            do {
                let imageData = try Data(contentsOf: imageFile)
                let success = try await sheet.embedImageAutoSized(
                    imageData,
                    at: coordinate,
                    of: workbook,
                    scale: 0.5
                )
                if success {
                    print("[INFO] ✓ Successfully embedded image at \(coordinate)")
                } else {
                    print("[WARNING] Failed to embed image at \(coordinate)")
                }
            } catch {
                print("[ERROR] Failed to embed image \(imageFile.lastPathComponent): \(error)")
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

    /// Calculates approximate text width for column sizing (8 pixels per character)
    private static func calculateTextWidth(_ text: String) -> Double {
        return Double(text.count) * 8.0
    }

    /// Returns sorted list of image file URLs from directory (PNG, JPG, JPEG, GIF)
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