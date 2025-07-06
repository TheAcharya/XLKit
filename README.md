# XLKit

**Excel spreadsheet (XLSX) format writer on pure SWIFT.**

XLKit is a library focused creating Excel spreadsheet (XLSX) files directly on mobile devices.


## Generation example XLSX file

```swift
    // library generate huge XLSX file with three sheets.
    XWorkBook.test()
``` 

## NEWS

- Add support inserting images into cells.

```swift
    //new kind value of XValue for icons/images 
    XValue.icon(XImageCell)
```
 
### example of inserting image

```swift
   let cell = sheet.AddCell(XCoords(row: 1, col: 1))
   let ImageCellValue:XImageCell = XImageCell(key: XImages.append(with: logoicon!)!, size: CGSize(width: 200, height: 75))
   // CGSize(width: 200, height: 75) - size display image in pixel
    cell.value = .icon(ImageCellValue)
```

- fix bug creating empty document

## Requirements

**Apple Platforms**

- Xcode 11.3 or later
- Swift 5.5 or later
- iOS 10.0

## Installation

### Swift Package Manager

[Swift Package Manager](https://swift.org/package-manager/) is a tool for
managing the distribution of Swift code. It’s integrated with the Swift build
system to automate the process of downloading, compiling, and linking
dependencies on all platforms.

Once you have your Swift package set up, adding `XLKit` as a dependency is as
easy as adding it to the `dependencies` value of your `Package.swift`.

```swift
dependencies: [
  .package(url: "https://github.com/3973770/XLKit", .upToNextMajor(from: "0.1.0"))
]
```

If you're using XLKit in an app built with Xcode, you can also add it as a direct
dependency [using Xcode's
GUI](https://developer.apple.com/documentation/xcode/adding_package_dependencies_to_your_app).

### Dependency

**SSZipArchive**
ZipArchive is a simple utility class for zipping and unzipping files on iOS, macOS and tvOS.
https://github.com/ZipArchive/ZipArchive.git


## Screenshots
![screenshot of invoce](https://raw.githubusercontent.com/3973770/XLKit/main/Sources/Sample/screen1.png)

![screenshot of icon test](https://raw.githubusercontent.com/3973770/XLKit/main/Sources/Sample/screen2.png)

![screenshot of performance test](https://raw.githubusercontent.com/3973770/XLKit/main/Sources/Sample/screen3.png) 

## Contributing

### Sponsorship

If this library saved you any amount of time or money, please consider [sponsoring
the work of its maintainer](https://github.com/sponsors/3973770). 
Any amount is appreciated and helps in maintaining the project.


## Example for use

```swift
   let book = XWorkBook()

   let color:[UIColor] = [.darkGray, .green, .lightGray, .orange, .systemPink, .cyan, .purple, .magenta, .blue]
   let colortext:[UIColor] = [.darkGray, .black, .white, .darkText, .lightText]


   func GetRandomFont() -> XFontName {
       let cases = XFontName.allCases
       return cases[Int.random(in: 0..<cases.count)]
   }


   var sheet = book.NewSheet("Invoice")

   var cell = sheet.AddCell(XCoords(row: 2, col: 6))
   cell.Cols(txt: .white, bg: .systemOrange)
   cell.value = .text("INVOICE")
   cell.Font = XFont(.TrebuchetMS, 16,true)
   cell.alignmentHorizontal = .center

   cell = sheet.AddCell(XCoords(row: 3, col: 6))
   cell.Cols(txt: .white, bg: .systemOrange)
   cell.value = .text("#12345")
   cell.Font = XFont(.TrebuchetMS, 12,true)
   cell.alignmentHorizontal = .left

   cell = sheet.AddCell(XCoords(row: 2, col: 1))
   cell.value = .text("Your company name")
   cell.Font = XFont(.TrebuchetMS, 16,true)

   cell = sheet.AddCell(XCoords(row: 3, col: 1))
   cell.value = .text("[Address Line 1]")

   cell = sheet.AddCell(XCoords(row: 4, col: 1))
   cell.value = .text("[Address Line 2]")

   cell = sheet.AddCell(XCoords(row: 5, col: 1))
   cell.value = .text("[Address Line 3]")

   cell = sheet.AddCell(XCoords(row: 7, col: 1))
   cell.color = .systemOrange
   cell.value = .text("Bill To:")
   cell.Font = XFont(.TrebuchetMS, 12,true)

   cell = sheet.AddCell(XCoords(row: 8, col: 1))
   cell.value = .text("[Address Line 1]")

   cell = sheet.AddCell(XCoords(row: 9, col: 1))
   cell.value = .text("[Address Line 2]")

   cell = sheet.AddCell(XCoords(row: 10, col: 1))
   cell.value = .text("[Address Line 3]")

   /// date
   cell = sheet.AddCell(XCoords(row: 13, col: 1))
   cell.color = .systemOrange
   cell.value = .text("Invoice Date")
   cell.Font = XFont(.TrebuchetMS, 12,true)

   cell = sheet.AddCell(XCoords(row: 14, col: 1))
   cell.value = .text("01/22/2022")

   /// term
   cell = sheet.AddCell(XCoords(row: 13, col: 2))

   cell.value = .text("Terms")
   cell.Font = XFont(.TrebuchetMS, 12,true)

   cell = sheet.AddCell(XCoords(row: 14, col: 2))
   cell.value = .text("30 days")

   /// Due Date
   cell = sheet.AddCell(XCoords(row: 13, col: 3))
   cell.color = .systemOrange
   cell.value = .text("Due Date")
   cell.Font = XFont(.TrebuchetMS, 12,true)

   cell = sheet.AddCell(XCoords(row: 14, col: 3))
   cell.value = .text("02/20/2022")

   /// table
   /// title
   var line = 16
   cell = sheet.AddCell(XCoords(row: line, col: 1))
   cell.Cols(txt: .white, bg: .systemOrange)
   cell.Border = true
   cell.value = .text("Description")

   cell = sheet.AddCell(XCoords(row: line, col: 4))
   cell.Cols(txt: .white, bg: .systemOrange)
   cell.Border = true
   cell.value = .text("Qty")
   cell.Font = XFont(.TrebuchetMS, 10,true)

   cell = sheet.AddCell(XCoords(row: line, col: 5))
   cell.Cols(txt: .white, bg: .systemOrange)
   cell.Border = true
   cell.value = .text("Unit Price")
   cell.Font = XFont(.TrebuchetMS, 10,true)

   cell = sheet.AddCell(XCoords(row: line, col: 6))
   cell.Cols(txt: .white, bg: .systemOrange)
   cell.Border = true
   cell.value = .text("Amount")
   cell.Font = XFont(.TrebuchetMS, 10,true)

   /// line
   line += 1
   cell = sheet.AddCell(XCoords(row: line, col: 1))
   cell.Border = true
   cell.value = .text("item #1")

   cell = sheet.AddCell(XCoords(row: line, col: 4))
   cell.Border = true
   cell.value = .double(3)
   cell.alignmentHorizontal = .right

   cell = sheet.AddCell(XCoords(row: line, col: 5))
   cell.Border = true
   cell.value = .double(50)
   cell.alignmentHorizontal = .right

   cell = sheet.AddCell(XCoords(row: line, col: 6))
   cell.Border = true
   cell.value = .double(150)
   cell.alignmentHorizontal = .right

   /// line
   line += 1
   cell = sheet.AddCell(XCoords(row: line, col: 1))
   if #available(iOS 13.0, *) {
       cell.Cols(txt: .black, bg: .systemGray6)
   } else {
       cell.Cols(txt: .black, bg: .lightGray)
   }
   cell.Border = true
   cell.value = .text("item #2")

   cell = sheet.AddCell(XCoords(row: line, col: 4))
   if #available(iOS 13.0, *) {
       cell.Cols(txt: .black, bg: .systemGray6)
   } else {
       cell.Cols(txt: .black, bg: .lightGray)
   }
   cell.Border = true
   cell.value = .double(2)
   cell.alignmentHorizontal = .right

   cell = sheet.AddCell(XCoords(row: line, col: 5))
   if #available(iOS 13.0, *) {
       cell.Cols(txt: .black, bg: .systemGray6)
   } else {
       cell.Cols(txt: .black, bg: .lightGray)
   }
   cell.Border = true
   cell.value = .double(100)
   cell.alignmentHorizontal = .right

   cell = sheet.AddCell(XCoords(row: line, col: 6))
   if #available(iOS 13.0, *) {
       cell.Cols(txt: .black, bg: .systemGray6)
   } else {
       cell.Cols(txt: .black, bg: .lightGray)
   }
   cell.Border = true
   cell.value = .double(200)
   cell.alignmentHorizontal = .right

   /// line
   line += 1
   cell = sheet.AddCell(XCoords(row: line, col: 1))
   cell.Border = true
   cell.value = .text("item #3")

   cell = sheet.AddCell(XCoords(row: line, col: 4))
   cell.Border = true
   cell.value = .double(4)
   cell.alignmentHorizontal = .right

   cell = sheet.AddCell(XCoords(row: line, col: 5))
   cell.Border = true
   cell.value = .double(200)
   cell.alignmentHorizontal = .right

   cell = sheet.AddCell(XCoords(row: line, col: 6))
   cell.Border = true
   cell.value = .double(800)
   cell.alignmentHorizontal = .right

   line += 2

   cell = sheet.AddCell(XCoords(row: line, col: 1))
   cell.value = .text("Thank you for your business!")
   cell.Font = XFont(.TrebuchetMS, 10, false, true)
   cell.alignmentHorizontal = .left

   cell = sheet.AddCell(XCoords(row: line, col: 5))
   cell.Cols(txt: .white, bg: .systemOrange)
   cell.value = .text("Total")
   cell.Font = XFont(.TrebuchetMS, 11, true)
   cell.alignmentHorizontal = .left

   cell = sheet.AddCell(XCoords(row: line, col: 6))
   cell.Cols(txt: .white, bg: .systemOrange)
   cell.value = .double(1100)
   cell.Font = XFont(.TrebuchetMS, 11, true)
   cell.alignmentHorizontal = .right

   line += 2

   cell = sheet.AddCell(XCoords(row: line, col: 1))
   cell.value = .text("Payment Options")
   cell.Font = XFont(.TrebuchetMS, 12, true, false)
   cell.alignmentHorizontal = .left

   line += 1

   cell = sheet.AddCell(XCoords(row: line, col: 1))
   cell.value = .text("Enter PayPal email address or bank account number here")
   cell.Font = XFont(.TrebuchetMS, 10, false, true)
   cell.alignmentHorizontal = .left

   sheet.buildindex()

   sheet.MergeRect(XRect(2, 1, 5, 1))
   sheet.MergeRect(XRect(3, 1, 5, 1))
   sheet.MergeRect(XRect(4, 1, 5, 1))
   sheet.MergeRect(XRect(5, 1, 5, 1))

   sheet.MergeRect(XRect(16, 1, 3, 1))
   sheet.MergeRect(XRect(17, 1, 3, 1))
   sheet.MergeRect(XRect(18, 1, 3, 1))
   sheet.MergeRect(XRect(19, 1, 3, 1))

   sheet.MergeRect(XRect(23, 1, 3, 1))
   sheet.MergeRect(XRect(24, 1, 3, 1))


   sheet.ForColumnSetWidth(1, 100)
   sheet.ForColumnSetWidth(2, 70)
   sheet.ForColumnSetWidth(3, 60)
   sheet.ForColumnSetWidth(4, 70)
   sheet.ForColumnSetWidth(5, 70)
   sheet.ForColumnSetWidth(6, 90)


   sheet = book.NewSheet("Perfomance1 Sheet")
   for col in 1...20 {
       sheet.ForColumnSetWidth(col,Int.random(in: 50..<100))
       for row in 1...1000 {
           let cell = sheet.AddCell(XCoords(row: row, col: col))
           cell.value = .integer(Int.random(in: 100..<200000))
           cell.Font = XFont(GetRandomFont(), Int.random(in: 10..<20))
           cell.color = colortext[Int.random(in: 0..<colortext.count)]
           cell.colorbackground = color[Int.random(in: 0..<color.count)]
           cell.Border = true
           cell.alignmentHorizontal = .center
           cell.alignmentVertical = .center
       }
   }

   sheet = book.NewSheet("Perfomance2 Sheet")
   for col in 1...20 {
       sheet.ForColumnSetWidth(col,Int.random(in: 50..<100))
       for row in 1...1000 {
           let cell = sheet.AddCell(XCoords(row: row, col: col))
           cell.value = .text("\(row):\(col)")
           cell.Font = XFont(GetRandomFont(), Int.random(in: 10..<20))
           cell.color = colortext[Int.random(in: 0..<colortext.count)]
           cell.colorbackground = color[Int.random(in: 0..<color.count)]
           cell.Border = true
           cell.alignmentHorizontal = .center
           cell.alignmentVertical = .center
       }
   }

   let fileid = book.save("example.xlsx")
   print("<<<File XLSX generated!>>>")
   print("\(fileid)")
``` 

## License

SSZipArchive is protected under the [MIT license](https://github.com/samsoffes/ssziparchive/raw/master/LICENSE) and our slightly modified version of [minizip-ng (formally minizip)](https://github.com/zlib-ng/minizip-ng) 3.0.2 is licensed under the [Zlib license](https://www.zlib.net/zlib_license.html).

## Acknowledgments

* Big thanks to [ZipArchive](https://github.com/ZipArchive).
