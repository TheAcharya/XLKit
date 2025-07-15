//
//  ImageSizingUtils.swift
//  XLKit • https://github.com/TheAcharya/XLKit
//  © 2025 Vigneswaran Rajkumar • Licensed under MIT License
//

import Foundation
import CoreGraphics

public struct ImageSizingUtils {
    // MARK: - Excel Constants
    
    /// 1 pixel = 9525 EMUs (Excel's internal coordinate system)
    private static let pixelsToEMUs: Int64 = 9525
    
    /// Column width conversion: pixels / 8.0 (derived from manual Excel analysis)
    private static let columnWidthFactor: CGFloat = 8.0
    
    /// Row height conversion: pixels / 1.33 (derived from manual Excel analysis)
    private static let rowHeightFactor: CGFloat = 1.33
    
    /// Calculates display size maintaining aspect ratio within max bounds
    /// - Parameter scale: Scaling factor (default: 0.5 for compact images)
    public static func calculateDisplaySize(
        originalSize: CGSize,
        maxWidth: CGFloat,
        maxHeight: CGFloat,
        minWidth: CGFloat = 100,
        minHeight: CGFloat = 60,
        scale: CGFloat = 0.5
    ) -> CGSize {
        let scaledMaxWidth = maxWidth * scale
        let scaledMaxHeight = maxHeight * scale
        
        let widthScale = scaledMaxWidth / originalSize.width
        let heightScale = scaledMaxHeight / originalSize.height
        
        // Use smaller scale to fit within bounds and preserve aspect ratio
        let maxScale = min(widthScale, heightScale, 1.0)
        
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

    /// Excel column width: pixels / 8.0
    public static func excelColumnWidth(forPixelWidth pixelWidth: CGFloat) -> CGFloat {
        return max(pixelWidth / columnWidthFactor, 0.0)
    }

    /// Excel row height: pixels / 1.33
    public static func excelRowHeight(forPixelHeight pixelHeight: CGFloat) -> CGFloat {
        return pixelHeight / rowHeightFactor
    }
    
    // MARK: - EMU Calculations
    
    /// Converts pixels to EMUs (Excel's internal format)
    public static func pixelsToEMUs(_ pixels: CGFloat) -> Int64 {
        return Int64(pixels * CGFloat(pixelsToEMUs))
    }
    
    /// Converts EMUs to pixels
    public static func emusToPixels(_ emus: Int64) -> CGFloat {
        return CGFloat(emus) / CGFloat(pixelsToEMUs)
    }
    
    /// Drawing dimensions in EMUs for Excel (maintains aspect ratio)
    public static func calculateDrawingDimensions(
        imageSize: CGSize,
        maxWidth: CGFloat = 800,
        maxHeight: CGFloat = 600
    ) -> (width: Int64, height: Int64) {
        let displaySize = calculateDisplaySize(
            originalSize: imageSize,
            maxWidth: maxWidth,
            maxHeight: maxHeight
        )
        
        let widthEMUs = pixelsToEMUs(displaySize.width)
        let heightEMUs = pixelsToEMUs(displaySize.height)
        
        return (width: widthEMUs, height: heightEMUs)
    }
    
    // MARK: - Positioning Calculations
    
    /// Positioning offsets for perfect image placement
    public static func calculatePositioningOffsets(
        imageSize: CGSize,
        cellWidth: CGFloat,
        cellHeight: CGFloat
    ) -> (x: Int64, y: Int64) {
        let cellWidthEMUs = pixelsToEMUs(cellWidth * columnWidthFactor)
        let cellHeightEMUs = pixelsToEMUs(cellHeight * rowHeightFactor)
        
        let xOffset = (cellWidthEMUs - pixelsToEMUs(imageSize.width)) / 2
        let yOffset = (cellHeightEMUs - pixelsToEMUs(imageSize.height)) / 2
        
        return (x: max(0, xOffset), y: max(0, yOffset))
    }
    
    /// Row offset for proper image positioning (3175 EMUs ≈ 0.33 pixels)
    public static func calculateRowOffset(imageHeight: CGFloat) -> Int64 {
        return 3175
    }

    /// Ideal Excel column width and row height to fit image exactly
    public static func idealCellSizeForImage(
        imageWidth: CGFloat,
        imageHeight: CGFloat
    ) -> (colWidth: CGFloat, rowHeight: CGFloat) {
        let colWidth = imageWidth / columnWidthFactor
        let rowHeight = imageHeight / rowHeightFactor
        return (colWidth, rowHeight)
    }

    /// Pixel size of cell given Excel column width and row height
    public static func cellPixelSize(
        colWidth: CGFloat,
        rowHeight: CGFloat
    ) -> (width: CGFloat, height: CGFloat) {
        let width = colWidth * columnWidthFactor
        let height = rowHeight * rowHeightFactor
        return (width, height)
    }

    /// Offsets to center image in cell (in EMUs)
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

    /// Drawing size in EMUs for given image dimensions
    public static func drawingEMUs(
        imageWidth: CGFloat,
        imageHeight: CGFloat
    ) -> (cx: Int64, cy: Int64) {
        return (pixelsToEMUs(imageWidth), pixelsToEMUs(imageHeight))
    }
} 