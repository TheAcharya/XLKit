import Foundation

final public class XSheet{
    
    private static let ABC = Array("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    public var title:String = ""
    private var cells:[XCell] = []
    var mergecells : [String] = []
    var merge : [XRect] = []
    var RowH : [Int:Int] = [:]
    var ColW : [Int:Int] = [:]
    private var indexcells : [String:XCell] = [:]
    var xml : String?
    var fix:XCoords = XCoords()
    var drawingsxml:String?
    var drawingsxmlrels:String?
    var drawingsSheetrels:String?
    
    public init() {}
    
    public init(_ title:String){
        self.title = title
    }
    
    public func ForRowSetHeight(_ row:Int,_ Height:Int){
        self.RowH[row] = Height
    }
    public func ForColumnSetWidth(_ column:Int,_ Width:Int){
        self.ColW[column] = Width
    }
    
    public var GetMaxRowCol : (row:Int,col:Int) {
        var maxRC:(row: Int,col: Int) = (row:0,col:0)
        for Cell in self.cells {
            guard let cellcoords = Cell.coords else {continue}
            if maxRC.row < cellcoords.row {
                maxRC.row = cellcoords.row
            }
            if maxRC.col < cellcoords.col {
                maxRC.col = cellcoords.col
            }
        }
        return maxRC
    }
    
    static func EncodeNumberABC(_ num:Int) -> String {
        guard num >= XSheet.ABC.count else { return "\(XSheet.ABC[num])"}
 
        let whole:Int = num / XSheet.ABC.count
        let remain:Int = num - whole*XSheet.ABC.count

        return "\(XSheet.EncodeNumberABC(whole-1))\(XSheet.EncodeNumberABC(remain))"
    }
    
    public func MergeRect(_ rect:XRect) {
        let xcord = XSheet.EncodeNumberABC(rect.col-1)
        let ycord = "\(rect.row)"
        let x2cord = XSheet.EncodeNumberABC(rect.col+rect.width-2)
        let y2cord = "\(rect.row+rect.height-1)"
        self.merge.append(rect)
        self.mergecells.append("\(xcord)\(ycord):\(x2cord)\(y2cord)")
        
        if let maincell = self.Get(XCoords(row: rect.row, col: rect.col)) {
            for row in rect.row...rect.row+rect.height-1 {
                for col in rect.col...rect.col+rect.width-1 {
                    if col == rect.col, row == rect.row {
                        continue
                    }
                    
                    let cord = XCoords(row: row, col: col)
                    if self.Get(cord) == nil {
                        let cellnew = self.AddCell(cord)
                        cellnew.value = .text("")
                        cellnew.Border = maincell.Border
                        cellnew.Font = maincell.Font
                        cellnew.color = maincell.color
                        cellnew.colorbackground = maincell.colorbackground
                    }
                }
            }
        }
    }
    
    /// add cell with coords to current sheet
    public func AddCell(_ coords:XCoords) -> XCell{
        if let cell = self.Get(coords) {
            return cell
        }else{
            let cellnew = XCell(coords)
            self.cells.append(cellnew)
            AddCellToIndex(cellnew)
            return cellnew
        }
    }
    
    /// append exestit cell to current sheet
    func append(_ newElement: XCell){
        self.cells.append(newElement)
        AddCellToIndex(newElement)
    }
    
    func AddCellToIndex(_ cell:XCell){
        if self.indexcells[cell.coords!.address] == nil {
            self.indexcells[cell.coords!.address] = cell
        }
    }
    
    /// build indez of all cell in sheet
    public func buildindex(){
        self.indexcells.removeAll()
        for cell in self.cells {
            self.indexcells[cell.coords!.address] = cell
        }
    }
    
    /// get cell from sheer by coords
    public func Get(_ coords:XCoords) -> XCell? {
        guard let ret = self.indexcells[coords.address] else { return nil }
        return ret
    }
    
    func GetMaxWidth(_ col:Int,_ numrows:Int) -> Int {
        var maxw=50;
        for row in 1...numrows {
            let cell = self.Get(XCoords(row: row, col: col))
            if cell != nil, !cell!.nocalculatewidth, maxw < cell!.width {
                maxw = cell!.width
            }
        }
        return maxw
    }
    
}

extension XSheet: Sequence {
    public func makeIterator() -> Array<XCell>.Iterator {
        return cells.makeIterator()
    }
    
}

fileprivate extension String{
    
    func XSheetTitle() -> String
    {
        let syms:[String] = [":","\\","/","?","*","[","]","     ","    ","   ","  "]
        var str = "\(self)"
        for sym in syms {
            str = str.replacingOccurrences(of: sym, with: " ", options: NSString.CompareOptions.literal, range: nil)
        }
        return str
    }
}
