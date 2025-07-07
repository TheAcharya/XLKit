//
//  ImageSizingUtils.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import CoreGraphics

public struct ImageSizingUtils {
    // MARK: - Excel Constants
    
    /// Excel uses EMUs (English Metric Units) for drawing dimensions
    /// 1 pixel = 9525 EMUs
    private static let pixelsToEMUs: Int64 = 9525
    
    /// Excel column width conversion factor
    /// Based on analysis of manually created Excel files
    /// Manual file uses width="102" for 816 pixels: 816 / 102 = 8.0
    private static let columnWidthFactor: CGFloat = 8.0
    
    /// Excel row height conversion factor (points)
    /// Based on analysis of manually created Excel files
    /// Manual file uses ht="344" for 459 pixels: 459 / 344 = 1.33
    private static let rowHeightFactor: CGFloat = 1.33
    
    /// Calculates the display size for an image, maintaining aspect ratio and fitting within max bounds.
    /// Always uses the same scale factor for both width and height to preserve aspect ratio 100%.
    /// - Parameter scale: Scaling factor for image size (default: 0.5 = 50% of max bounds for more compact images)
    public static func calculateDisplaySize(
        originalSize: CGSize,
        maxWidth: CGFloat,
        maxHeight: CGFloat,
        minWidth: CGFloat = 100,
        minHeight: CGFloat = 60,
        scale: CGFloat = 0.5 // Default 50% scaling for more compact images
    ) -> CGSize {
        let scaledMaxWidth = maxWidth * scale
        let scaledMaxHeight = maxHeight * scale
        
        // Calculate scale factors for both dimensions
        let widthScale = scaledMaxWidth / originalSize.width
        let heightScale = scaledMaxHeight / originalSize.height
        
        // Use the smaller scale factor to ensure the image fits within bounds
        // This guarantees aspect ratio preservation
        let maxScale = min(widthScale, heightScale, 1.0) // Don't scale up, only down
        
        // Apply the same scale to both dimensions
        let displayWidth = originalSize.width * maxScale
        let displayHeight = originalSize.height * maxScale
        
        // Ensure minimum size for very small images
        if displayWidth < minWidth || displayHeight < minHeight {
            let minWidthScale = minWidth / originalSize.width
            let minHeightScale = minHeight / originalSize.height
            let minScale = max(minWidthScale, minHeightScale)
            return CGSize(
                width: originalSize.width * minScale,
                height: originalSize.height * minScale
            )
        }
        
        return CGSize(width: displayWidth, height: displayHeight)
    }

    /// Calculates the Excel column width for a given pixel width
    /// Based on analysis of manually created Excel files that preserve aspect ratio
    public static func excelColumnWidth(forPixelWidth pixelWidth: CGFloat) -> CGFloat {
        // Excel column width formula: width = pixels / 8.0
        // This formula is derived from analyzing manually created Excel files
        // that correctly preserve image aspect ratios
        return max(pixelWidth / columnWidthFactor, 0.0)
    }

    /// Calculates the Excel row height in points for a given pixel height
    /// Based on analysis of manually created Excel files that preserve aspect ratio
    public static func excelRowHeight(forPixelHeight pixelHeight: CGFloat) -> CGFloat {
        // Excel row height formula: height = pixels / 1.33
        // This formula is derived from analyzing manually created Excel files
        // that correctly preserve image aspect ratios
        return pixelHeight / rowHeightFactor
    }
    
    // MARK: - EMU Calculations (Excel's Internal Format)
    
    /// Converts pixels to EMUs (English Metric Units) for Excel drawing dimensions
    public static func pixelsToEMUs(_ pixels: CGFloat) -> Int64 {
        return Int64(pixels * CGFloat(pixelsToEMUs))
    }
    
    /// Converts EMUs to pixels
    public static func emusToPixels(_ emus: Int64) -> CGFloat {
        return CGFloat(emus) / CGFloat(pixelsToEMUs)
    }
    
    /// Calculates drawing dimensions in EMUs for Excel
    /// This is what Excel actually uses internally for perfect aspect ratio preservation
    public static func calculateDrawingDimensions(
        imageSize: CGSize,
        maxWidth: CGFloat = 800,
        maxHeight: CGFloat = 600
    ) -> (width: Int64, height: Int64) {
        // Calculate display size maintaining aspect ratio
        let displaySize = calculateDisplaySize(
            originalSize: imageSize,
            maxWidth: maxWidth,
            maxHeight: maxHeight
        )
        
        // Convert to EMUs (Excel's internal format)
        let widthEMUs = pixelsToEMUs(displaySize.width)
        let heightEMUs = pixelsToEMUs(displaySize.height)
        
        return (width: widthEMUs, height: heightEMUs)
    }
    
    // MARK: - Positioning Calculations
    
    /// Calculates positioning offsets for perfect image placement
    /// Based on analysis of manually created Excel files
    public static func calculatePositioningOffsets(
        imageSize: CGSize,
        cellWidth: CGFloat,
        cellHeight: CGFloat
    ) -> (x: Int64, y: Int64) {
        // Convert cell dimensions to EMUs
        let cellWidthEMUs = pixelsToEMUs(cellWidth * columnWidthFactor)
        let cellHeightEMUs = pixelsToEMUs(cellHeight * rowHeightFactor)
        
        // Calculate offsets to center the image within the cell
        let xOffset = (cellWidthEMUs - pixelsToEMUs(imageSize.width)) / 2
        let yOffset = (cellHeightEMUs - pixelsToEMUs(imageSize.height)) / 2
        
        return (x: max(0, xOffset), y: max(0, yOffset))
    }
    
    /// Calculates row offset for proper image positioning
    /// Based on analysis of manually created Excel files
    public static func calculateRowOffset(imageHeight: CGFloat) -> Int64 {
        // Manual file uses rowOff="3175" for images
        // This is approximately 3175 EMUs = 0.33 pixels
        // We'll use a small offset to ensure proper positioning
        return 3175
    }

    /// Given an image size in pixels, returns the ideal Excel column width and row height (in Excel units) to fit the image exactly.
    public static func idealCellSizeForImage(
        imageWidth: CGFloat,
        imageHeight: CGFloat
    ) -> (colWidth: CGFloat, rowHeight: CGFloat) {
        // Use consistent formulas derived from manual Excel file analysis
        // Column width: pixels / 8.0
        // Row height: pixels / 1.33
        let colWidth = imageWidth / columnWidthFactor
        let rowHeight = imageHeight / rowHeightFactor
        return (colWidth, rowHeight)
    }

    /// Given an image size and a cell size (in Excel units), returns the pixel size of the cell.
    public static func cellPixelSize(
        colWidth: CGFloat,
        rowHeight: CGFloat
    ) -> (width: CGFloat, height: CGFloat) {
        // Excel column width to pixels: pixels = colWidth * 8.0
        // Excel row height to pixels: pixels = rowHeight * 1.33
        let width = colWidth * columnWidthFactor
        let height = rowHeight * rowHeightFactor
        return (width, height)
    }

    /// Given an image and a cell, returns the offsets (in EMUs) to center the image in the cell.
    public static func imageOffsetsInCell(
        imageWidth: CGFloat,
        imageHeight: CGFloat,
        cellWidth: CGFloat,
        cellHeight: CGFloat
    ) -> (xEMU: Int64, yEMU: Int64) {
        let x = max(0, (cellWidth - imageWidth) / 2)
        let y = max(0, (cellHeight - imageHeight) / 2)
        return (pixelsToEMUs(x), pixelsToEMUs(y))
    }

    /// Given an image size, returns the drawing size in EMUs.
    public static func drawingEMUs(
        imageWidth: CGFloat,
        imageHeight: CGFloat
    ) -> (cx: Int64, cy: Int64) {
        return (pixelsToEMUs(imageWidth), pixelsToEMUs(imageHeight))
    }
} 