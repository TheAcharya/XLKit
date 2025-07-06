#if os(macOS)
    import Cocoa
    public typealias FontClass  = NSFont
    public typealias ColorClass = NSColor
    public typealias ImageClass = NSImage
#else
    import UIKit
    public typealias FontClass  = UIFont
    public typealias ColorClass = UIColor
    public typealias ImageClass = UIImage
#endif

public struct XLKit {
    static func test() -> Bool {
        XWorkBook.test()
    }
}
