//
//  XLSXEngine.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
@preconcurrency import XLKitCore
import XLKitFormatters
import XLKitImages
import ZIPFoundation

/// XLSX file generation and XML/ZIP utilities for XLKit
// (Content will be filled in next step) 

// MARK: - XLSX/ZIP/XML Engine for XLKit

/// XLSX file generation and XML/ZIP utilities for XLKit
public struct XLSXEngine {
    
    /// Generates XLSX file asynchronously
    @MainActor
    public static func generateXLSX(workbook: Workbook, to url: URL) throws {
        // Security checks
        try SecurityManager.checkRateLimit()
        try CoreUtils.validateFilePath(url.path)
        
        SecurityManager.logSecurityOperation("xlsx_generation_started", details: [
            "target_path": url.path,
            "workbook_sheets": workbook.getSheets().count,
            "workbook_images": workbook.imageCount
        ])
        let tempDir = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        
        defer {
            try? FileManager.default.removeItem(at: tempDir)
        }
        
        // Create required directories
        let xlDir = tempDir.appendingPathComponent("xl")
        let worksheetsDir = xlDir.appendingPathComponent("worksheets")
        let themeDir = xlDir.appendingPathComponent("theme")
        let mediaDir = xlDir.appendingPathComponent("media")
        let drawingsDir = xlDir.appendingPathComponent("drawings")
        let docPropsDir = tempDir.appendingPathComponent("docProps")
        let relsDir = tempDir.appendingPathComponent("_rels")
        let xlRelsDir = xlDir.appendingPathComponent("_rels")
        let worksheetsRelsDir = worksheetsDir.appendingPathComponent("_rels")
        let drawingsRelsDir = drawingsDir.appendingPathComponent("_rels")
        
        try FileManager.default.createDirectory(at: xlDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: worksheetsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: themeDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: mediaDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: drawingsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: docPropsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: relsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: xlRelsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: worksheetsRelsDir, withIntermediateDirectories: true)
        try FileManager.default.createDirectory(at: drawingsRelsDir, withIntermediateDirectories: true)
        
        // Generate Content_Types.xml
        let contentTypes = generateContentTypes()
        try contentTypes.write(to: tempDir.appendingPathComponent("[Content_Types].xml"), atomically: true, encoding: .utf8)
        
        // Generate docProps files
        let (appXml, coreXml) = try generateDocProps()
        try appXml.write(to: docPropsDir.appendingPathComponent("app.xml"), atomically: true, encoding: .utf8)
        try coreXml.write(to: docPropsDir.appendingPathComponent("core.xml"), atomically: true, encoding: .utf8)
        
        // Generate theme
        let themeXml = generateTheme()
        try themeXml.write(to: themeDir.appendingPathComponent("theme1.xml"), atomically: true, encoding: .utf8)
        
        // Generate styles and shared strings
        let (formatMapping, sharedStrings) = try generateStylesAndStrings(xlDir: xlDir, workbook: workbook)
        
        // Generate workbook
        try generateWorkbook(xlDir: xlDir, workbook: workbook)
        
        // Generate media files and drawings
        try generateMediaAndDrawings(mediaDir: mediaDir, drawingsDir: drawingsDir, workbook: workbook)
        
        // Generate worksheets
        try generateWorksheets(worksheetsDir: worksheetsDir, workbook: workbook, formatMapping: formatMapping, sharedStrings: sharedStrings)
        
        // Generate relationships
        try generateRelationships(tempDir: tempDir, xlDir: xlDir, worksheetsDir: worksheetsDir, drawingsDir: drawingsDir, workbook: workbook)
        
        // Create ZIP archive
        try createZIPArchive(from: tempDir, to: url)
        
        // Generate and store checksum
        let checksum = try CoreUtils.generateFileChecksum(url)
        SecurityManager.storeFileChecksum(checksum, for: url)
        
        SecurityManager.logSecurityOperation("xlsx_generation_completed", details: [
            "target_path": url.path,
            "checksum": checksum,
            "file_size": try Data(contentsOf: url).count
        ])
    }
    

    
    // MARK: - Private XLSX Generation Methods
    

    
    private static func generateWorkbook(xlDir: URL, workbook: Workbook) throws {
        let content = generateWorkbookXML(workbook: workbook)
        try content.write(to: xlDir.appendingPathComponent("workbook.xml"), atomically: true, encoding: .utf8)
    }
    
    private static func generateWorkbookXML(workbook: Workbook) -> String {
        var content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        content += "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        content += "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>"
        content += "<workbookPr defaultThemeVersion=\"124226\"/>"
        content += "<bookViews>"
        content += "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>"
        content += "</bookViews>"
        content += "<sheets>"
        
        for sheet in workbook.getSheets() {
            content += "<sheet name=\"\(CoreUtils.escapeXML(sheet.name))\" sheetId=\"\(sheet.id)\" r:id=\"rId\(sheet.id)\"/>"
        }
        
        content += "</sheets>"
        content += "<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
        content += "</workbook>"
        
        return content
    }
    
    private static func generateStylesAndStrings(xlDir: URL, workbook: Workbook) throws -> ([String: Int], [String: Int]) {
        // Collect all unique formats from all sheets
        var uniqueFormats: [CellFormat] = []
        var formatToId: [String: Int] = [:]
        var stringToId: [String: Int] = [:]
        var uniqueNumberFormats: [String] = []
        var numberFormatToId: [String: Int] = [:]
        
        // Collect all unique strings, formats, and number formats
        for sheet in workbook.getSheets() {
            for coordinate in sheet.getUsedCells() {
                if let value = sheet.getCell(coordinate) {
                    let stringValue = value.stringValue
                    if stringToId[stringValue] == nil {
                        stringToId[stringValue] = stringToId.count
                    }
                }
                
                if let format = sheet.getCellFormat(coordinate) {
                    let formatKey = formatToKey(format)
                    if formatToId[formatKey] == nil {
                        formatToId[formatKey] = uniqueFormats.count + 1 // Start from 1, 0 is default
                        uniqueFormats.append(format)
                    }
                    
                    // Collect unique number formats
                    if let numberFormat = format.numberFormat {
                        let numberFormatString = numberFormat == .custom ? (format.customNumberFormat ?? "") : numberFormat.rawValue
                        if !uniqueNumberFormats.contains(numberFormatString) {
                            uniqueNumberFormats.append(numberFormatString)
                            numberFormatToId[numberFormatString] = numberFormatToId.count + 164 // Start from 164 (Excel's custom format range)
                        }
                    }
                }
            }
        }
        
        // Generate shared strings XML
        var sharedStringsContent = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        sharedStringsContent += "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"\(stringToId.count)\" uniqueCount=\"\(stringToId.count)\">"
        
        for (stringValue, _) in stringToId.sorted(by: { $0.value < $1.value }) {
            sharedStringsContent += "<si><t>\(CoreUtils.escapeXML(stringValue))</t></si>"
        }
        
        sharedStringsContent += "</sst>"
        
        try sharedStringsContent.write(to: xlDir.appendingPathComponent("sharedStrings.xml"), atomically: true, encoding: .utf8)
        
        // Generate styles XML
        var content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        content += "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        
        // Number formats section
        content += "<numFmts count=\"\(uniqueNumberFormats.count + 1)\">"
        content += "<numFmt numFmtId=\"0\" formatCode=\"General\"/>"
        
        for numberFormatString in uniqueNumberFormats {
            let numFmtId = numberFormatToId[numberFormatString] ?? 164
            content += "<numFmt numFmtId=\"\(numFmtId)\" formatCode=\"\(CoreUtils.escapeXML(numberFormatString))\"/>"
        }
        
        content += "</numFmts>"
        
        // Fonts
        content += "<fonts count=\"\(uniqueFormats.count + 1)\">"
        content += "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>"
        
        // Generate font definitions for each unique format
        for format in uniqueFormats {
            content += "<font>"
            
            if let fontWeight = format.fontWeight {
                if fontWeight == .bold {
                    content += "<b/>"
                }
            }
            
            if let fontSize = format.fontSize {
                content += "<sz val=\"\(fontSize)\"/>"
            } else {
                content += "<sz val=\"11\"/>"
            }
            
            // Handle font color - use custom color if specified, otherwise use theme
            if let fontColor = format.fontColor {
                content += "<color rgb=\"\(fontColor.replacingOccurrences(of: "#", with: ""))\"/>"
            } else {
                content += "<color theme=\"1\"/>"
            }
            
            if let fontName = format.fontName {
                content += "<name val=\"\(fontName)\"/>"
            } else {
                content += "<name val=\"Calibri\"/>"
            }
            
            content += "<family val=\"2\"/><scheme val=\"minor\"/></font>"
        }
        
        content += "</fonts>"
        
        // Fills
        content += "<fills count=\"\(uniqueFormats.count + 2)\">"
        content += "<fill><patternFill patternType=\"none\"/></fill>"
        content += "<fill><patternFill patternType=\"gray125\"/></fill>"
        
        // Generate fill definitions for each unique format
        for format in uniqueFormats {
            content += "<fill>"
            
            if let backgroundColor = format.backgroundColor {
                content += "<patternFill patternType=\"solid\"><fgColor rgb=\"\(backgroundColor.replacingOccurrences(of: "#", with: ""))\"/></patternFill>"
            } else {
                content += "<patternFill patternType=\"none\"/>"
            }
            
            content += "</fill>"
        }
        
        content += "</fills>"
        
        // Borders
        content += "<borders count=\"1\">"
        content += "<border><left/><right/><top/><bottom/><diagonal/></border>"
        content += "</borders>"
        
        // Cell style formats
        content += "<cellStyleXfs count=\"1\">"
        content += "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
        content += "</cellStyleXfs>"
        
        // Cell formats
        content += "<cellXfs count=\"\(uniqueFormats.count + 1)\">"
        content += "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        
        // Generate cell format definitions
        for (index, format) in uniqueFormats.enumerated() {
            let fontId = index + 1
            let fillId = index + 2
            
            // Determine number format ID
            var numFmtId = 0 // Default to General
            if let numberFormat = format.numberFormat {
                let numberFormatString = numberFormat == .custom ? (format.customNumberFormat ?? "") : numberFormat.rawValue
                numFmtId = numberFormatToId[numberFormatString] ?? 0
            }
            
            var xf = "<xf numFmtId=\"\(numFmtId)\" fontId=\"\(fontId)\" fillId=\"\(fillId)\" borderId=\"0\" xfId=\"0\""
            
            var applyFont = false
            if format.fontWeight == .bold || format.fontStyle != nil || format.fontName != nil || format.fontSize != nil || format.fontColor != nil {
                applyFont = true
            }
            if applyFont {
                xf += " applyFont=\"1\""
            }
            
            // Apply number format if specified
            if format.numberFormat != nil {
                xf += " applyNumberFormat=\"1\""
            }
            
            if format.horizontalAlignment != nil || format.verticalAlignment != nil {
                xf += " applyAlignment=\"1\">"
                xf += "<alignment"
                if let horizontalAlignment = format.horizontalAlignment {
                    xf += " horizontal=\"\(horizontalAlignment.rawValue)\""
                }
                if let verticalAlignment = format.verticalAlignment {
                    xf += " vertical=\"\(verticalAlignment.rawValue)\""
                }
                xf += "/>"
                xf += "</xf>"
            } else {
                xf += "/>"
            }
            
            content += xf
        }
        
        content += "</cellXfs>"
        
        // Cell styles
        content += "<cellStyles count=\"1\">"
        content += "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
        content += "</cellStyles>"
        
        // Differential formats
        content += "<dxfs count=\"0\"/>"
        
        // Table styles
        content += "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>"
        
        content += "</styleSheet>"
        
        try content.write(to: xlDir.appendingPathComponent("styles.xml"), atomically: true, encoding: .utf8)
        
        return (formatToId, stringToId)
    }
    
    private static func generateWorksheets(worksheetsDir: URL, workbook: Workbook, formatMapping: [String: Int], sharedStrings: [String: Int]) throws {
        for sheet in workbook.getSheets() {
            let content = generateWorksheetXML(sheet: sheet, formatMapping: formatMapping, sharedStrings: sharedStrings)
            try content.write(to: worksheetsDir.appendingPathComponent("sheet\(sheet.id).xml"), atomically: true, encoding: .utf8)
            
            // Generate worksheet relationships if there are images
            if !sheet.getImages().isEmpty {
                try generateWorksheetRels(worksheetRelsDir: worksheetsDir.appendingPathComponent("_rels"), sheet: sheet, formatMapping: formatMapping)
            }
        }
    }
    
    private static func generateWorksheetRels(worksheetRelsDir: URL, sheet: Sheet, formatMapping: [String: Int]) throws {
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
    
    private static func generateWorksheetXML(sheet: Sheet, formatMapping: [String: Int], sharedStrings: [String: Int]) -> String {
        var content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        content += "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        
        // Add dimension
        let maxRow = sheet.getUsedCells().compactMap { CellCoordinate(excelAddress: $0) }.map { $0.row }.max() ?? 1
        let maxCol = sheet.getUsedCells().compactMap { CellCoordinate(excelAddress: $0) }.map { $0.column }.max() ?? 1
        let maxColLetter = CoreUtils.columnLetter(from: maxCol)
        content += "<dimension ref=\"A1:\(maxColLetter)\(maxRow)\"/>"
        
        // Add sheet views
        content += "<sheetViews>"
        content += "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
        content += "</sheetViews>"
        
        // Add sheet format properties
        content += "<sheetFormatPr defaultRowHeight=\"15\"/>"
        
        // Add column widths if any
        if !sheet.getColumnWidths().isEmpty {
            content += generateColumnWidthsXML(sheet: sheet)
        }
        
        // Add row heights if any
        if !sheet.getRowHeights().isEmpty {
            content += generateRowHeightsXML(sheet: sheet)
        }
        
        content += "<sheetData>"
        
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
            let rowCells = rows[rowNum] ?? [:]
            let minCol = rowCells.keys.compactMap { CellCoordinate(excelAddress: $0) }.map { $0.column }.min() ?? 1
            let maxCol = rowCells.keys.compactMap { CellCoordinate(excelAddress: $0) }.map { $0.column }.max() ?? 1
            
            // Check for custom row height
            if let customHeight = sheet.getRowHeight(rowNum) {
                content += "<row r=\"\(rowNum)\" spans=\"\(minCol):\(maxCol)\" ht=\"\(customHeight)\" customHeight=\"1\">"
            } else {
                content += "<row r=\"\(rowNum)\" spans=\"\(minCol):\(maxCol)\">"
            }
            
            for coordinate in rowCells.keys.sorted() {
                guard let value = rowCells[coordinate] else { continue }
                let format = sheet.getCellFormat(coordinate)
                content += generateCellXML(coordinate: coordinate, value: value, format: format, formatMapping: formatMapping, sharedStrings: sharedStrings)
            }
            
            content += "</row>"
        }
        
        content += "</sheetData>"
        
        // Add page margins
        content += "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
        
        // Add drawing reference if sheet has images
        let sheetImages = sheet.getImages()
        if !sheetImages.isEmpty {
            content += "<drawing r:id=\"rId1\"/>"
        }
        
        content += "</worksheet>"
        
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
        var content = ""
        
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
    
    private static func generateCellXML(coordinate: String, value: CellValue, format: CellFormat?, formatMapping: [String: Int], sharedStrings: [String: Int]) -> String {
        let styleId = format != nil ? getStyleId(for: format!, formatMapping: formatMapping) : nil
        let styleAttribute = styleId != nil ? " s=\"\(styleId!)\"" : ""
        
        switch value {
        case .string(let stringValue):
            // Find the shared string ID using the stringToId mapping
            let stringId = sharedStrings[stringValue] ?? 0
            return """
            <c r="\(coordinate)"\(styleAttribute) t="s"><v>\(stringId)</v></c>
            """
        case .number(let numberValue):
            return """
            <c r="\(coordinate)"\(styleAttribute) t="n"><v>\(numberValue)</v></c>
            """
        case .integer(let intValue):
            return """
            <c r="\(coordinate)"\(styleAttribute) t="n"><v>\(intValue)</v></c>
            """
        case .boolean(let boolValue):
            return """
            <c r="\(coordinate)"\(styleAttribute) t="b"><v>\(boolValue ? 1 : 0)</v></c>
            """
        case .date(let dateValue):
            let excelNumber = CoreUtils.excelNumberFromDate(dateValue)
            return """
            <c r="\(coordinate)"\(styleAttribute) t="n"><v>\(excelNumber)</v></c>
            """
        case .formula(let formulaValue):
            return """
            <c r="\(coordinate)"\(styleAttribute) t="str"><f>\(CoreUtils.escapeXML(formulaValue))</f></c>
            """
        case .empty:
            return """
            <c r="\(coordinate)"\(styleAttribute)/>
            """
        }
    }
    
    public static func formatToKey(_ format: CellFormat) -> String {
        var key = ""
        key += "fontName:\(format.fontName ?? "nil")"
        key += "fontSize:\(format.fontSize ?? 0)"
        key += "fontWeight:\(format.fontWeight?.rawValue ?? "nil")"
        key += "fontStyle:\(format.fontStyle?.rawValue ?? "nil")"
        key += "fontColor:\(format.fontColor ?? "nil")"
        key += "backgroundColor:\(format.backgroundColor ?? "nil")"
        key += "horizontalAlignment:\(format.horizontalAlignment?.rawValue ?? "nil")"
        key += "verticalAlignment:\(format.verticalAlignment?.rawValue ?? "nil")"
        
        // Include number format information
        if let numberFormat = format.numberFormat {
            let numberFormatString = numberFormat == .custom ? (format.customNumberFormat ?? "") : numberFormat.rawValue
            key += "numberFormat:\(numberFormatString)"
        } else {
            key += "numberFormat:nil"
        }
        
        return key
    }
    
    private static func getStyleId(for format: CellFormat, formatMapping: [String: Int]) -> Int? {
        // Use the global format mapping that was created during styles.xml generation
        let key = formatToKey(format)
        return formatMapping[key]
    }
    
    private static func generateMediaAndDrawings(mediaDir: URL, drawingsDir: URL, workbook: Workbook) throws {
        // Generate media files
        try generateMediaFiles(mediaDir: mediaDir, workbook: workbook)
        
        // Generate drawings for each sheet that has images
        for sheet in workbook.getSheets() {
            let sheetImages = sheet.getImages()
            if !sheetImages.isEmpty {
                try generateDrawingXML(drawingsDir: drawingsDir, sheet: sheet, workbook: workbook)
            }
        }
    }
    
    private static func generateMediaFiles(mediaDir: URL, workbook: Workbook) throws {
        for image in workbook.getImages() {
            let imageURL = mediaDir.appendingPathComponent("\(image.id).\(image.format.rawValue)")
            try image.data.write(to: imageURL)
        }
    }
    
    private static func generateDrawingXML(drawingsDir: URL, sheet: Sheet, workbook: Workbook) throws {
        let drawingId = "drawing\(sheet.id)"
        let drawingXML = generateDrawingXMLContent(sheet: sheet, workbook: workbook)
        try drawingXML.write(to: drawingsDir.appendingPathComponent("\(drawingId).xml"), atomically: true, encoding: .utf8)
        
        // Generate drawing relationships
        try generateDrawingRelationships(drawingsDir: drawingsDir, drawingId: drawingId, sheet: sheet, workbook: workbook)
    }
    
    private static func generateDrawingXMLContent(sheet: Sheet, workbook: Workbook) -> String {
        var content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        content += "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
        
        let sheetImages = sheet.getImages()
        var imageIndex = 1
        
        for (coordinate, image) in sheetImages {
            guard let cellCoord = CellCoordinate(excelAddress: coordinate) else { continue }
            let col = cellCoord.column
            let row = cellCoord.row

            // Use the image's display size if available, else original size
            let imgSize = image.displaySize ?? image.originalSize

            // 1. Compute ideal cell size (Excel units)
            let (idealColWidth, idealRowHeight) = ImageSizingUtils.idealCellSizeForImage(
                imageWidth: imgSize.width,
                imageHeight: imgSize.height
            )
            // 2. Compute cell pixel size
            let (cellPixelWidth, cellPixelHeight) = ImageSizingUtils.cellPixelSize(
                colWidth: idealColWidth,
                rowHeight: idealRowHeight
            )
            // 3. Compute drawing size in EMUs
            let (cx, cy) = ImageSizingUtils.drawingEMUs(
                imageWidth: imgSize.width,
                imageHeight: imgSize.height
            )
            // 4. Compute offsets to center image in cell
            let (offsetX, offsetY) = ImageSizingUtils.imageOffsetsInCell(
                imageWidth: imgSize.width,
                imageHeight: imgSize.height,
                cellWidth: cellPixelWidth,
                cellHeight: cellPixelHeight
            )
            // 5. Set cell size in the sheet (so Excel renders the cell big enough)
            sheet.setColumnWidth(col, width: idealColWidth)
            sheet.setRowHeight(row, height: idealRowHeight)
            // 6. Use rowOff=0 (Excel default) for now
            let rowOff = 0

            content += "<xdr:twoCellAnchor editAs=\"oneCell\">"
            content += "<xdr:from><xdr:col>\(col - 1)</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>\(row - 1)</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            content += "<xdr:to><xdr:col>\(col)</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>\(row)</xdr:row><xdr:rowOff>\(rowOff)</xdr:rowOff></xdr:to>"
            content += "<xdr:pic>"
            content += "<xdr:nvPicPr>"
            content += "<xdr:cNvPr id=\"\(imageIndex * 2)\" name=\"Picture \(imageIndex)\">"
            content += "<a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\">"
            content += "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{\(UUID().uuidString.uppercased())}\"/>"
            content += "</a:ext></a:extLst>"
            content += "</xdr:cNvPr>"
            content += "<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr>"
            content += "</xdr:nvPicPr>"
            content += "<xdr:blipFill>"
            content += "<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId\(imageIndex)\" cstate=\"print\">"
            content += "<a:extLst><a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">"
            content += "<a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/>"
            content += "</a:ext></a:extLst>"
            content += "</a:blip>"
            content += "<a:stretch><a:fillRect/></a:stretch>"
            content += "</xdr:blipFill>"
            content += "<xdr:spPr>"
            content += "<a:xfrm><a:off x=\"\(offsetX)\" y=\"\(offsetY)\"/><a:ext cx=\"\(cx)\" cy=\"\(cy)\"/></a:xfrm>"
            content += "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
            content += "</xdr:spPr>"
            content += "</xdr:pic>"
            content += "<xdr:clientData/>"
            content += "</xdr:twoCellAnchor>"
            imageIndex += 1
        }
        
        content += "</xdr:wsDr>"
        return content
    }
    
    private static func generateDrawingRelationships(drawingsDir: URL, drawingId: String, sheet: Sheet, workbook: Workbook) throws {
        var content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        content += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        
        let sheetImages = sheet.getImages()
        var imageIndex = 1
        
        for (_, image) in sheetImages {
            content += "<Relationship Id=\"rId\(imageIndex)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/\(image.id).\(image.format.rawValue)\"/>"
            imageIndex += 1
        }
        
        content += "</Relationships>"
        
        try content.write(to: drawingsDir.appendingPathComponent("_rels/\(drawingId).xml.rels"), atomically: true, encoding: .utf8)
    }
    
    // MARK: - ZIP Archive Creation
    
    private static func createZIPArchive(from sourceDir: URL, to destinationURL: URL) throws {
        // Remove existing file if it exists
        if FileManager.default.fileExists(atPath: destinationURL.path) {
            try FileManager.default.removeItem(at: destinationURL)
        }
        
        // Use ZIPFoundation for cross-platform compatibility and security
        // Create a temporary ZIP file first, then move it to the destination
        let tempZipURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".zip")
        
        do {
            // Create ZIP archive with contents of sourceDir (not the directory itself)
            let archive = try Archive(url: tempZipURL, accessMode: .create)
            
            let parentDir = sourceDir.deletingLastPathComponent()
            let enumerator = FileManager.default.enumerator(at: parentDir, includingPropertiesForKeys: [.isDirectoryKey])
            
            // Normalize the source directory path to handle /private prefix
            let normalizedSourcePath = sourceDir.resolvingSymlinksInPath().path
            let prefix = normalizedSourcePath + "/"
            
            while let fileURL = enumerator?.nextObject() as? URL {
                // Normalize the file path to handle /private prefix
                let normalizedFilePath = fileURL.resolvingSymlinksInPath().path
                
                // Only add files inside sourceDir
                guard normalizedFilePath.hasPrefix(prefix) else { continue }
                
                let relativePath = String(normalizedFilePath.dropFirst(prefix.count))
                if relativePath.isEmpty { continue }
                
                let resourceValues = try fileURL.resourceValues(forKeys: [.isDirectoryKey])
                if resourceValues.isDirectory == true { continue }
                
                try archive.addEntry(with: relativePath, fileURL: fileURL)
            }
            
            // On iOS, we need to handle file system restrictions differently
            #if os(iOS)
            // For iOS, try to copy the file instead of moving it
            // This works better with iOS sandbox restrictions
            try FileManager.default.copyItem(at: tempZipURL, to: destinationURL)
            try FileManager.default.removeItem(at: tempZipURL)
            #else
            // On macOS, we can move the file directly
            try FileManager.default.moveItem(at: tempZipURL, to: destinationURL)
            #endif
            
        } catch {
            try? FileManager.default.removeItem(at: tempZipURL)
            throw XLKitError.zipCreationError("ZIP creation failed: \(error.localizedDescription)")
        }
    }
    
    private static func generateRelationships(tempDir: URL, xlDir: URL, worksheetsDir: URL, drawingsDir: URL, workbook: Workbook) throws {
        // Root relationships
        let rootRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>"
        try rootRels.write(to: tempDir.appendingPathComponent("_rels/.rels"), atomically: true, encoding: .utf8)
        
        // Generate workbook relationships dynamically
        var workbookRelsContent = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        
        // Add worksheet relationships
        for sheet in workbook.getSheets() {
            workbookRelsContent += "<Relationship Id=\"rId\(sheet.id)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet\(sheet.id).xml\"/>"
        }
        
        // Add other workbook relationships
        let nextId = workbook.getSheets().count + 1
        workbookRelsContent += "<Relationship Id=\"rId\(nextId)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
        workbookRelsContent += "<Relationship Id=\"rId\(nextId + 1)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
        workbookRelsContent += "<Relationship Id=\"rId\(nextId + 2)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
        workbookRelsContent += "</Relationships>"
        
        try workbookRelsContent.write(to: xlDir.appendingPathComponent("_rels/workbook.xml.rels"), atomically: true, encoding: .utf8)
        
        // Worksheet relationships - check if sheet has images
        for sheet in workbook.getSheets() {
            let sheetImages = sheet.getImages()
            if !sheetImages.isEmpty {
                let drawingId = "drawing\(sheet.id)"
                let worksheetRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/\(drawingId).xml\"/></Relationships>"
                try worksheetRels.write(to: worksheetsDir.appendingPathComponent("_rels/sheet\(sheet.id).xml.rels"), atomically: true, encoding: .utf8)
            } else {
                let worksheetRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>"
                try worksheetRels.write(to: worksheetsDir.appendingPathComponent("_rels/sheet\(sheet.id).xml.rels"), atomically: true, encoding: .utf8)
            }
        }
    }
    
    private static func generateContentTypes() -> String {
        var content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        content += "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        content += "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        content += "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        content += "<Default Extension=\"png\" ContentType=\"image/png\"/>"
        content += "<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>"
        content += "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>"
        content += "<Default Extension=\"gif\" ContentType=\"image/gif\"/>"
        content += "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
        content += "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
        content += "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
        content += "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>"
        content += "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        content += "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        content += "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
        content += "<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>"
        content += "</Types>"
        return content
    }
    
    private static func generateDocProps() throws -> (String, String) {
        // Generate app.xml
        let appXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=\"1\" baseType=\"lpstr\"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><Pages>1</Pages><Words>0</Words><Characters>0</Characters><PresentationFormat></PresentationFormat><Paragraphs>0</Paragraphs><Slides>0</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>"
        
        // Generate core.xml
        let coreXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>XLKit</dc:creator><cp:lastModifiedBy>XLKit</cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">2025-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2025-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>"
        
        return (appXml, coreXml)
    }
    
    private static func generateTheme() -> String {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2><a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera><a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig></a:scene3d><a:sp3d><a:spPr><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:satMod val=\"350000\"/><a:shade val=\"99000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path></a:gradFill></a:spPr></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:satMod val=\"350000\"/><a:shade val=\"99000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"350000\"/><a:tint val=\"80000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"300000\"/><a:shade val=\"30000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E424DB5}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-B589-74238A4823A8}\"/></a:ext></a:extLst></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
    }
}

 