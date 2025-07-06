import Foundation
#if os(macOS)
    import Cocoa
#else
    import UIKit
#endif

// Mark: struct coordinate of cell
public struct XCoords {
    public var row:Int = 0
    public var col:Int = 0
    private var caddr:String?
    public var address: String {
        if let addr = self.caddr {
            return addr
        }else{
            return "\(self.col):\(self.row)"
        }
    }
    
    init (){}
    public init (row:Int,col:Int){
        self.row = row
        self.col = col
        self.caddr = "\(self.col):\(self.row)"
    }
}

public struct XRect {
    public var row:Int = 0
    public var col:Int = 0
    public var width:Int = 0
    public var height:Int = 0
    public init(_ row:Int,_ col:Int,_ width:Int,_ height:Int){
        self.row = row
        self.col = col
        self.width = width
        self.height = height
    }
    init (){}
    
    public func inrect(_ coord:XCoords) -> Bool{
        if self.row <= coord.row, self.col <= coord.col, coord.col <= self.col+self.width-1, coord.row <= self.row+self.height-1 {
            if self.row == coord.row, self.col == coord.col {
                return false
            }else{
                return true
            }
        }else{
            return false
        }
    }
}

public struct XFont {
    public var Font:XFontName
    public var FontSize:Int
    public var bold:Bool
    public var italic:Bool
    public var strike:Bool
    public var underline:Bool
    public init (_ Font:XFontName,_ FontSize:Int,_ bold:Bool=false,_ italic:Bool=false,_ strike:Bool=false,_ underline:Bool=false){
        self.Font = Font
        self.FontSize = FontSize
        self.bold = bold
        self.italic = italic
        self.strike = strike
        self.underline = underline
    }
    
    public var getfont:FontClass?
    {
        if let font = FontClass(name: self.Font.rawValue, size: CGFloat(self.FontSize)) {
            return font
        }else{
            return FontClass(name: "Arial", size: CGFloat(self.FontSize))
        }
    }
}

public enum XValue : Equatable {
    case long(UInt64)
    case integer(Int)
    case text(String)
    case double(Double)
    case float(Float)
    case icon(XImageCell)
}


public enum XAligmentHorizontal:UInt64 {
    case left, center, right
    
    public func str() -> String {
        switch self {
            case .left:
                return "left"
            case .center:
                return "center"
            case .right:
                return "right"
        }
    }
    public func id() -> UInt64 {
        switch self {
            case .left:
                return 1
            case .center:
                return 2
            case .right:
                return 3
        }
    }
}

public enum XAligmentVertical:UInt64 {
    case top, center, bottom
    
    func str() -> String {
        switch self {
            case .top:
                return "top"
            case .center:
                return "center"
            case .bottom:
                return "bottom"
        }
    }
    func id() -> UInt64 {
        switch self {
            case .top:
                return 1
            case .center:
                return 2
            case .bottom:
                return 3
        }
    }
}

final public class XCell{
    public var coords:XCoords?
    public var value:XValue?
    public var Font:XFont?
    public var alignmentVertical:XAligmentVertical = .center
    public var alignmentHorizontal:XAligmentHorizontal = .left
    public var color : ColorClass = .black
    public var colorbackground : ColorClass = .white

    public var width:Int = 50
    
    public var Border:Bool = false
    
    var idFont:Int  = 0
    var idFill:Int  = 0
    var idStyle:Int = 0
    var idVal:Int?
    var nocalculatewidth: Bool = true

    public init(_ coords:XCoords){
        self.coords = coords
        self.Font = XFont(.TrebuchetMS, 10)
    }
    
    public func Cols(txt color: ColorClass,bg bgcolor: ColorClass){
        self.color = color
        self.colorbackground = bgcolor
    }

    public func Als(v Vertical:XAligmentVertical,h Horizontal:XAligmentHorizontal){
        self.alignmentVertical = Vertical
        self.alignmentHorizontal = Horizontal
    }
}
