//
//  XLSXEngine.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
@preconcurrency import XLKitCore
import XLKitFormatters
import XLKitImages

/// XLSX file generation and XML/ZIP utilities for XLKit
// (Content will be filled in next step) 

// MARK: - XLSX/ZIP/XML Engine for XLKit

/// XLSX file generation and XML/ZIP utilities for XLKit
public struct XLSXEngine {
    
    /// Generates XLSX file asynchronously
    public static func generateXLSX(workbook: Workbook, to url: URL) async throws {
        try await withCheckedThrowingContinuation { continuation in
            DispatchQueue.global(qos: .userInitiated).async {
                do {
                    try generateXLSXSync(workbook: workbook, to: url)
                    continuation.resume()
                } catch {
                    continuation.resume(throwing: error)
                }
            }
        }
    }
    
    /// Generates XLSX file synchronously
    public static func generateXLSXSync(workbook: Workbook, to url: URL) throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent("xlkit-\(UUID().uuidString)")
        
        do {
            try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
            
            // Create required directories
            let xlDir = tempDir.appendingPathComponent("xl")
            let worksheetsDir = xlDir.appendingPathComponent("worksheets")
            let mediaDir = xlDir.appendingPathComponent("media")
            let relsDir = tempDir.appendingPathComponent("_rels")
            let xlRelsDir = xlDir.appendingPathComponent("_rels")
            let worksheetRelsDir = worksheetsDir.appendingPathComponent("_rels")
            
            try FileManager.default.createDirectory(at: xlDir, withIntermediateDirectories: true)
            try FileManager.default.createDirectory(at: worksheetsDir, withIntermediateDirectories: true)
            try FileManager.default.createDirectory(at: mediaDir, withIntermediateDirectories: true)
            try FileManager.default.createDirectory(at: relsDir, withIntermediateDirectories: true)
            try FileManager.default.createDirectory(at: xlRelsDir, withIntermediateDirectories: true)
            try FileManager.default.createDirectory(at: worksheetRelsDir, withIntermediateDirectories: true)
            
            // Generate content files
            try generateContentTypes(tempDir: tempDir, workbook: workbook)
            try generateRels(tempDir: tempDir, workbook: workbook)
            try generateWorkbookRels(xlRelsDir: xlRelsDir, workbook: workbook)
            try generateWorkbook(xlDir: xlDir, workbook: workbook)
            try generateWorksheets(worksheetsDir: worksheetsDir, workbook: workbook)
            try generateMediaFiles(mediaDir: mediaDir, workbook: workbook)
            
            // Create ZIP archive
            try createZIPArchive(from: tempDir, to: url)
            
            // Clean up
            try FileManager.default.removeItem(at: tempDir)
            
        } catch {
            // Clean up on error
            try? FileManager.default.removeItem(at: tempDir)
            throw error
        }
    }
    
    // MARK: - Private XLSX Generation Methods
    
    private static func generateContentTypes(tempDir: URL, workbook: Workbook) throws {
        var content = """
        <?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
        <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
            <Default Extension=\"xml\" ContentType=\"application/xml\"/>
            <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
            <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
        """
        
        for sheet in workbook.getSheets() {
            content += """
            
            <Override PartName=\"/xl/worksheets/sheet\(sheet.id).xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
            """
        }
        
        // Add image content types
        for image in workbook.getImages() {
            content += """
            
            <Override PartName=\"/xl/media/\(image.id).\(image.format.rawValue)\" ContentType=\"\(image.format.excelContentType)\"/>
            """
        }
        
        content += """
        
        </Types>
        """
        
        try content.write(to: tempDir.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)
    }
    
    private static func generateRels(tempDir: URL, workbook: Workbook) throws {
        let content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """
        
        try content.write(to: tempDir.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorkbookRels(xlRelsDir: URL, workbook: Workbook) throws {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        """
        
        for sheet in workbook.getSheets() {
            content += """
            
            <Relationship Id="rId\(sheet.id)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet\(sheet.id).xml"/>
            """
        }
        
        content += """
        
        </Relationships>
        """
        
        try content.write(to: xlRelsDir.appendingPathComponent("workbook.xml.rels"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorkbook(xlDir: URL, workbook: Workbook) throws {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets>
        """
        
        for sheet in workbook.getSheets() {
            content += """
            
                <sheet name="\(CoreUtils.escapeXML(sheet.name))" sheetId="\(sheet.id)" r:id="rId\(sheet.id)"/>
            """
        }
        
        content += """
        
            </sheets>
        </workbook>
        """
        
        try content.write(to: xlDir.appendingPathComponent("workbook.xml"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorksheets(worksheetsDir: URL, workbook: Workbook) throws {
        for sheet in workbook.getSheets() {
            let content = generateWorksheetXML(sheet: sheet)
            try content.write(to: worksheetsDir.appendingPathComponent("sheet\(sheet.id).xml"), atomically: true, encoding: .utf8)
            
            // Generate worksheet relationships if there are images
            if !sheet.getImages().isEmpty {
                try generateWorksheetRels(worksheetRelsDir: worksheetsDir.appendingPathComponent("_rels"), sheet: sheet)
            }
        }
    }
    
    private static func generateWorksheetRels(worksheetRelsDir: URL, sheet: Sheet) throws {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        """
        
        var imageId = 1
        for (_, image) in sheet.getImages() {
            content += """
            
            <Relationship Id="rId\(imageId)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/\(image.id).\(image.format.rawValue)"/>
            """
            imageId += 1
        }
        
        content += """
        
        </Relationships>
        """
        
        try content.write(to: worksheetRelsDir.appendingPathComponent("sheet\(sheet.id).xml.rels"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorksheetXML(sheet: Sheet) -> String {
        var content = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        """
        
        // Add column widths if any
        if !sheet.getColumnWidths().isEmpty {
            content += generateColumnWidthsXML(sheet: sheet)
        }
        
        // Add row heights if any
        if !sheet.getRowHeights().isEmpty {
            content += generateRowHeightsXML(sheet: sheet)
        }
        
        content += """
        
            <sheetData>
        """
        
        // Group cells by row
        var rows: [Int: [String: CellValue]] = [:]
        for (coordinate, value) in sheet.getUsedCells().compactMap({ ($0, sheet.getCell($0)) }) {
            guard let cellCoord = CellCoordinate(excelAddress: coordinate) else { continue }
            if rows[cellCoord.row] == nil {
                rows[cellCoord.row] = [:]
            }
            rows[cellCoord.row]?[coordinate] = value
        }
        
        // Generate row XML
        for rowNum in rows.keys.sorted() {
            content += """
            
                <row r="\(rowNum)">
            """
            
            for coordinate in rows[rowNum]?.keys.sorted() ?? [] {
                guard let value = rows[rowNum]?[coordinate] else { continue }
                content += generateCellXML(coordinate: coordinate, value: value)
            }
            
            content += """
            
                </row>
            """
        }
        
        content += """
        
            </sheetData>
        """
        
        // Add images if any
        if !sheet.getImages().isEmpty {
            content += generateDrawingXML(sheet: sheet)
        }
        
        content += """
        
        </worksheet>
        """
        
        return content
    }
    
    private static func generateColumnWidthsXML(sheet: Sheet) -> String {
        var content = """
        
            <cols>
        """
        
        for (column, width) in sheet.getColumnWidths().sorted(by: { $0.key < $1.key }) {
            content += """
            
                <col min="\(column)" max="\(column)" width="\(width)" customWidth="1"/>
            """
        }
        
        content += """
        
            </cols>
        """
        
        return content
    }
    
    private static func generateRowHeightsXML(sheet: Sheet) -> String {
        var content = """
        
            <sheetFormatPr defaultRowHeight="15"/>
        """
        
        for (row, height) in sheet.getRowHeights().sorted(by: { $0.key < $1.key }) {
            content += """
            
            <row r="\(row)" ht="\(height)" customHeight="1"/>
            """
        }
        
        return content
    }
    
    private static func generateDrawingXML(sheet: Sheet) -> String {
        let content = """
        
            <drawing r:id="rId1"/>
        """
        
        return content
    }
    
    private static func generateCellXML(coordinate: String, value: CellValue) -> String {
        switch value {
        case .string(let stringValue):
            return """
            
                    <c r="\(coordinate)" t="s">
                        <v>\(CoreUtils.escapeXML(stringValue))</v>
                    </c>
            """
        case .number(let numberValue):
            return """
            
                    <c r="\(coordinate)" t="n">
                        <v>\(numberValue)</v>
                    </c>
            """
        case .integer(let intValue):
            return """
            
                    <c r="\(coordinate)" t="n">
                        <v>\(intValue)</v>
                    </c>
            """
        case .boolean(let boolValue):
            return """
            
                    <c r="\(coordinate)" t="b">
                        <v>\(boolValue ? 1 : 0)</v>
                    </c>
            """
        case .date(let dateValue):
            let excelNumber = CoreUtils.excelNumberFromDate(dateValue)
            return """
            
                    <c r="\(coordinate)" t="n">
                        <v>\(excelNumber)</v>
                    </c>
            """
        case .formula(let formulaValue):
            return """
            
                    <c r="\(coordinate)" t="str">
                        <f>\(CoreUtils.escapeXML(formulaValue))</f>
                    </c>
            """
        case .empty:
            return """
            
                    <c r="\(coordinate)"/>
            """
        }
    }
    
    private static func generateMediaFiles(mediaDir: URL, workbook: Workbook) throws {
        for image in workbook.getImages() {
            let imageURL = mediaDir.appendingPathComponent("\(image.id).\(image.format.rawValue)")
            try image.data.write(to: imageURL)
        }
    }
    
    // MARK: - ZIP Archive Creation
    
    private static func createZIPArchive(from sourceDir: URL, to destinationURL: URL) throws {
        // Remove existing file if it exists
        if FileManager.default.fileExists(atPath: destinationURL.path) {
            try FileManager.default.removeItem(at: destinationURL)
        }
        
        // Use system zip command for simplicity
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/zip")
        process.arguments = ["-r", destinationURL.path, "."]
        process.currentDirectoryURL = sourceDir
        
        try process.run()
        process.waitUntilExit()
        
        if process.terminationStatus != 0 {
            throw XLKitError.zipCreationError("ZIP creation failed with status \(process.terminationStatus)")
        }
    }
}

// (Insert all public definitions for XLSX generation, ZIP, and XML helpers here, as previously defined in the monolithic code) 