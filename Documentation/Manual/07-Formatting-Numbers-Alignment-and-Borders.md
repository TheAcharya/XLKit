# Chapter 07 — Formatting, Number Formats, Alignment, and Borders

Navigation: [← Chapter 06](06-Images-Embedding-and-Sizing.md) · [Manual index](README.md) · [Chapter 08 →](08-Security-and-Validation.md)

This chapter covers cell appearance: fonts, colours, borders, number formats, and text alignment. Examples build on `Workbook`, `Sheet`, `Cell`, and `CellFormat` types from [Chapter 03](03-Core-Model-Workbook-Sheet-and-Cells.md).


### Cell Formatting

```swift
let sheet = workbook.addSheet(name: "Formatted")

// Use predefined formats with convenience methods (recommended)
sheet.setCell("A1", string: "Header", format: CellFormat.header())
sheet.setCell("B1", number: 1234.56, format: CellFormat.currency())
sheet.setCell("C1", number: 0.85, format: CellFormat.percentage())
sheet.setCell("D1", date: Date(), format: CellFormat.date())
sheet.setCell("E1", boolean: true, format: CellFormat.text())

// Use Cell struct for more control
sheet.setCell("F1", cell: Cell.string("Header", format: CellFormat.header()))
sheet.setCell("G1", cell: Cell.number(1234.56, format: CellFormat.currency()))

// Custom formatting (`CellFormat` uses `init()` then set properties)
var customFormat = CellFormat()
customFormat.fontName = "Arial"
customFormat.fontSize = 14.0
customFormat.fontWeight = .bold
customFormat.backgroundColor = "#E0E0E0"
customFormat.horizontalAlignment = .center
sheet.setCell("H1", string: "Custom", format: customFormat)

// Font colour formatting
let redTextFormat = CellFormat.text(color: "#FF0000")
let blueTextFormat = CellFormat.text(color: "#0000FF")
let currencyRedFormat = CellFormat.currency(color: "#FF0000")

sheet.setCell("I1", string: "Red Text", format: redTextFormat)
sheet.setCell("J1", string: "Blue Text", format: blueTextFormat)
sheet.setCell("K1", number: 1234.56, format: currencyRedFormat)

// Border formatting
var borderedFormat = CellFormat.bordered()
borderedFormat.borderTop = .thin
borderedFormat.borderBottom = .thin
borderedFormat.borderLeft = .thin
borderedFormat.borderRight = .thin
borderedFormat.borderColor = "#000000"
sheet.setCell("L1", string: "Bordered Cell", format: borderedFormat)

// Get cell with formatting
let cellWithFormat = sheet.getCellWithFormat("A1")
let cellFormat = sheet.getCellFormat("A1")
```

### Font Colour Support

XLKit provides comprehensive font colour support with proper XML generation and theme colour support:

```swift
// Basic font colours
let redText = CellFormat.text(color: "#FF0000")
let blueText = CellFormat.text(color: "#0000FF")
let greenText = CellFormat.text(color: "#00FF00")

// Font colours with other formatting
var boldRedText = CellFormat()
boldRedText.fontColor = "#FF0000"
boldRedText.fontWeight = .bold
boldRedText.fontSize = 14.0

let colouredCurrency = CellFormat.currency(
    format: .currencyWithDecimals,
    color: "#FF0000"  // Red currency values
)

// Apply font colours to cells
sheet.setCell("A1", string: "Red Header", format: redText)
sheet.setCell("B1", number: 1234.56, format: colouredCurrency)
sheet.setCell("C1", string: "Bold Blue", format: boldRedText)

// Font colours work with all cell types
sheet.setCell("D1", boolean: true, format: CellFormat.text(color: "#FF6600"))
sheet.setCell("E1", date: Date(), format: CellFormat.text(color: "#6600FF"))
```

Supported colour formats:
- Hex colours: `#FF0000`, `#00FF00`, `#0000FF`
- Theme colours: Excel's built-in theme colour system
- All colours are properly converted to Excel's internal format

### Number Format Support

XLKit provides comprehensive number formatting support with proper Excel compliance. All number formats are correctly applied in Excel with thousands grouping, currency symbols, and proper display in the "Format Cells" dialog.

### Supported Number Formats

```swift
// Currency formatting
let currencyFormat = CellFormat.currency()                    // $1,234.56
let customCurrency = CellFormat.currency(format: .custom("$#,##0.00"))  // Custom currency

// Percentage formatting  
let percentageFormat = CellFormat.percentage()                // 12.34%
let customPercentage = CellFormat.percentage(format: .custom("0.00%"))  // Custom percentage

// Date formatting
let dateFormat = CellFormat.date()                            // Standard date format
let customDate = CellFormat.date(format: .custom("mmmm dd, yyyy"))      // Custom date

// Custom number formats
let thousandsFormat = CellFormat()
thousandsFormat.numberFormat = .custom
thousandsFormat.customNumberFormat = "#,##0"                  // 1,234,567

let decimalFormat = CellFormat()
decimalFormat.numberFormat = .custom
decimalFormat.customNumberFormat = "0.000"                    // 3.142

let mixedFormat = CellFormat()
mixedFormat.numberFormat = .custom
mixedFormat.customNumberFormat = "$#,##0.00;($#,##0.00)"     // ($1,234.56) for negatives
```

### Number Format Examples

```swift
let sheet = workbook.addSheet(name: "Number Formats")

// Currency examples
sheet.setCell("A1", number: 1234.56, format: CellFormat.currency())
sheet.setCell("A2", number: 5678.90, format: CellFormat.currency(format: .custom("$#,##0.00")))

// Percentage examples
sheet.setCell("B1", number: 0.1234, format: CellFormat.percentage())
sheet.setCell("B2", number: 0.5678, format: CellFormat.percentage(format: .custom("0.00%")))

// Custom number formats
var thousandsFormat = CellFormat()
thousandsFormat.numberFormat = .custom
thousandsFormat.customNumberFormat = "#,##0"
sheet.setCell("C1", number: 1234567, format: thousandsFormat)

var decimalFormat = CellFormat()
decimalFormat.numberFormat = .custom
decimalFormat.customNumberFormat = "0.000"
sheet.setCell("C2", number: 3.14159, format: decimalFormat)

// Date formatting
sheet.setCell("D1", date: Date(), format: CellFormat.date())
var customDateFormat = CellFormat()
customDateFormat.numberFormat = .custom
customDateFormat.customNumberFormat = "mmmm dd, yyyy"
sheet.setCell("D2", date: Date(), format: customDateFormat)
```

### Font Colour with Number Formats

```swift
// Red currency values
let redCurrencyFormat = CellFormat.currency(
    format: .currencyWithDecimals,
    color: "#FF0000"
)

// Blue percentage values
let bluePercentageFormat = CellFormat.percentage(
    format: .percentageWithDecimals,
    color: "#0000FF"
)

// Apply coloured number formats
sheet.setCell("A1", number: 1234.56, format: redCurrencyFormat)
sheet.setCell("B1", number: 0.85, format: bluePercentageFormat)
```

### Excel Compliance

All number formats are properly implemented with:
- Thousands grouping: Numbers display with proper separators (1,234.56)
- Currency symbols: Dollar signs and other currency symbols display correctly
- Format Cells dialog: Excel shows the correct format instead of "General"
- Custom formats: User-defined number formats work as expected
- Negative numbers: Proper handling of negative values with parentheses or minus signs
- Locale support: Respects system locale settings for decimal separators

### Testing Number Formats

Use the XLKitTestRunner to validate number formatting:

```bash
swift run XLKitTestRunner number-formats
```

This generates `Number-Format-Test.xlsx` with comprehensive examples of all number format types.

## Text Alignment Support

XLKit provides comprehensive text alignment support with all 6 alignment options available in Excel:

#### Horizontal alignment (5 options)

`CellFormat` exposes `public init()`; set properties after construction:

```swift
var leftAligned = CellFormat()
leftAligned.horizontalAlignment = .left

var centerAligned = CellFormat()
centerAligned.horizontalAlignment = .center

var rightAligned = CellFormat()
rightAligned.horizontalAlignment = .right

var justified = CellFormat()
justified.horizontalAlignment = .justify

var distributed = CellFormat()
distributed.horizontalAlignment = .distributed
```

#### Vertical alignment (5 options)

```swift
var topAligned = CellFormat()
topAligned.verticalAlignment = .top

var centerVertically = CellFormat()
centerVertically.verticalAlignment = .center

var bottomAligned = CellFormat()
bottomAligned.verticalAlignment = .bottom

var justifiedVertically = CellFormat()
justifiedVertically.verticalAlignment = .justify

var distributedVertically = CellFormat()
distributedVertically.verticalAlignment = .distributed
```

#### Combined alignment

```swift
var centerCenter = CellFormat()
centerCenter.horizontalAlignment = .center
centerCenter.verticalAlignment = .center

var topRight = CellFormat()
topRight.horizontalAlignment = .right
topRight.verticalAlignment = .top

var bottomLeft = CellFormat()
bottomLeft.horizontalAlignment = .left
bottomLeft.verticalAlignment = .bottom
```

#### Practical examples

```swift
var headerFormat = CellFormat()
headerFormat.fontWeight = .bold
headerFormat.fontSize = 14.0
headerFormat.backgroundColor = "#E6E6E6"
headerFormat.horizontalAlignment = .center

var numberFormat = CellFormat()
numberFormat.numberFormat = .currencyWithDecimals
numberFormat.horizontalAlignment = .right

var multiLineFormat = CellFormat()
multiLineFormat.textWrapping = true
multiLineFormat.verticalAlignment = .top

sheet.setCell("A1", string: "Centered Header", format: headerFormat)
sheet.setCell("B1", number: 1234.56, format: numberFormat)
sheet.setCell("C1", string: "Multi-line\ntext content", format: multiLineFormat)
```

#### Predefined formats with alignment

```swift
let header = CellFormat.header(fontSize: 14.0, backgroundColor: "#E6E6E6")

var currencyRight = CellFormat.currency(
    format: .currencyWithDecimals,
    color: "#FF0000"
)
currencyRight.horizontalAlignment = .right
```

#### Alignment with other formatting

```swift
var boldRedCenter = CellFormat()
boldRedCenter.fontWeight = .bold
boldRedCenter.fontColor = "#FF0000"
boldRedCenter.horizontalAlignment = .center
boldRedCenter.verticalAlignment = .center

var borderedDistributed = CellFormat()
borderedDistributed.borderTop = .thin
borderedDistributed.borderBottom = .thin
borderedDistributed.borderLeft = .thin
borderedDistributed.borderRight = .thin
borderedDistributed.horizontalAlignment = .distributed
borderedDistributed.verticalAlignment = .distributed
```

All alignment options are properly converted to Excel's OpenXML format and will display correctly in Excel, Google Sheets, LibreOffice Calc, and other spreadsheet applications.

