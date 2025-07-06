import Foundation

public enum XFontName:String, CaseIterable{
    case AcademyEngravedLET = "AcademyEngravedLetPlain"
    case AlNile = "AlNile"
    case AmericanTypewriter = "AmericanTypewriter"
    case AppleColorEmoji = "AppleColorEmoji"
    case AppleSDGothicNeo = "AppleSDGothicNeo-Regular"
    case AppleSymbols = "AppleSymbols"
    case Arial = "ArialMT"
    case ArialHebrew = "ArialHebrew"
    case ArialRoundedMTBold = "ArialRoundedMTBold"
    case Avenir = "Avenir-Book"
    case AvenirNext = "AvenirNext-Regular"
    case AvenirNextCondensed = "AvenirNextCondensed-Regular"
    case Baskerville = "Baskerville"
    case Bodoni72 = "BodoniSvtyTwoITCTT-Book"
    case Bodoni72Oldstyle = "BodoniSvtyTwoOSITCTT-Book"
    case Bodoni72Smallcaps = "BodoniSvtyTwoSCITCTT-Book"
    case BodoniOrnaments = "BodoniOrnamentsITCTT"
    case BradleyHand = "BradleyHandITCTT-Bold"
    case ChalkboardSE = "ChalkboardSE-Regular"
    case Chalkduster = "Chalkduster"
    case Charter = "Charter-Roman"
    case Cochin = "Cochin"
    case Copperplate = "Copperplate"
    case CourierNew = "CourierNewPSMT"
    case Damascus = "Damascus"
    case DevanagariSangamMN = "DevanagariSangamMN"
    case Didot = "Didot"
    case DINAlternate = "DINAlternate-Bold"
    case DINCondensed = "DINCondensed-Bold"
    case EuphemiaUCAS = "EuphemiaUCAS"
    case Farah = "Farah"
    case Futura = "Futura-Medium"
    case Galvji = "Galvji"
    case GeezaPro = "GeezaPro"
    case Georgia = "Georgia"
    case GillSans = "GillSans"
    case GranthaSangamMN = "GranthaSangamMN-Regular"
    case Helvetica = "Helvetica"
    case HelveticaNeue = "HelveticaNeue"
    case HiraginoMaruGothicProN = "HiraMaruProN-W4"
    case HiraginoMinchoProN = "HiraMinProN-W3"
    case HiraginoSans = "HiraginoSans-W3"
    case HoeflerText = "HoeflerText-Regular"
    case Kailasa = "Kailasa"
    case Kefa = "Kefa-Regular"
    case KhmerSangamMN = "KhmerSangamMN"
    case KohinoorBangla = "KohinoorBangla-Regular"
    case KohinoorDevanagari = "KohinoorDevanagari-Regular"
    case KohinoorGujarati = "KohinoorGujarati-Regular"
    case KohinoorTelugu = "KohinoorTelugu-Regular"
    case LaoSangamMN = "LaoSangamMN"
    case MalayalamSangamMN = "MalayalamSangamMN"
    case MarkerFelt = "MarkerFelt-Thin"
    case Menlo = "Menlo-Regular"
    case Mishafi = "DiwanMishafi"
    case MuktaMahee = "MuktaMahee-Regular"
    case MyanmarSangamMN = "MyanmarSangamMN"
    case Noteworthy = "Noteworthy-Light"
    case NotoNastaliqUrdu = "NotoNastaliqUrdu"
    case NotoSansKannada = "NotoSansKannada-Regular"
    case NotoSansMyanmar = "NotoSansMyanmar-Regular"
    case NotoSansOriya = "NotoSansOriya"
    case Optima = "Optima-Regular"
    case Palatino = "Palatino-Roman"
    case Papyrus = "Papyrus"
    case PartyLET = "PartyLetPlain"
    case PingFangHK = "PingFangHK-Regular"
    case PingFangSC = "PingFangSC-Regular"
    case PingFangTC = "PingFangTC-Regular"
    case Rockwell = "Rockwell-Regular"
    case SavoyeLET = "SavoyeLetPlain"
    case SinhalaSangamMN = "SinhalaSangamMN"
    case SnellRoundhand = "SnellRoundhand"
    case Symbol = "Symbol"
    case TamilSangamMN = "TamilSangamMN"
    case Thonburi = "Thonburi"
    case TimesNewRoman = "TimesNewRomanPSMT"
    case TrebuchetMS = "TrebuchetMS"
    case Verdana = "Verdana"
    case ZapfDingbats = "ZapfDingbatsITC"
    case Zapfino = "Zapfino"

    func ind() -> UInt64 {
        switch self {
        case .AcademyEngravedLET:
             return 1
        case .AlNile:
             return 2
        case .AmericanTypewriter:
             return 3
        case .AppleColorEmoji:
             return 4
        case .AppleSDGothicNeo:
             return 5
        case .AppleSymbols:
             return 6
        case .Arial:
             return 7
        case .ArialHebrew:
             return 8
        case .ArialRoundedMTBold:
             return 9
        case .Avenir:
             return 10
        case .AvenirNext:
             return 11
        case .AvenirNextCondensed:
             return 12
        case .Baskerville:
             return 13
        case .Bodoni72:
             return 14
        case .Bodoni72Oldstyle:
             return 15
        case .Bodoni72Smallcaps:
             return 16
        case .BodoniOrnaments:
             return 17
        case .BradleyHand:
             return 18
        case .ChalkboardSE:
             return 19
        case .Chalkduster:
             return 20
        case .Charter:
             return 21
        case .Cochin:
             return 22
        case .Copperplate:
             return 23
        case .CourierNew:
             return 24
        case .Damascus:
             return 25
        case .DevanagariSangamMN:
             return 26
        case .Didot:
             return 27
        case .DINAlternate:
             return 28
        case .DINCondensed:
             return 29
        case .EuphemiaUCAS:
             return 30
        case .Farah:
             return 31
        case .Futura:
             return 32
        case .Galvji:
             return 33
        case .GeezaPro:
             return 34
        case .Georgia:
             return 35
        case .GillSans:
             return 36
        case .GranthaSangamMN:
             return 37
        case .Helvetica:
             return 38
        case .HelveticaNeue:
             return 39
        case .HiraginoMaruGothicProN:
             return 40
        case .HiraginoMinchoProN:
             return 41
        case .HiraginoSans:
             return 42
        case .HoeflerText:
             return 43
        case .Kailasa:
             return 44
        case .Kefa:
             return 45
        case .KhmerSangamMN:
             return 46
        case .KohinoorBangla:
             return 47
        case .KohinoorDevanagari:
             return 48
        case .KohinoorGujarati:
             return 49
        case .KohinoorTelugu:
             return 50
        case .LaoSangamMN:
             return 51
        case .MalayalamSangamMN:
             return 52
        case .MarkerFelt:
             return 53
        case .Menlo:
             return 54
        case .Mishafi:
             return 55
        case .MuktaMahee:
             return 56
        case .MyanmarSangamMN:
             return 57
        case .Noteworthy:
             return 58
        case .NotoNastaliqUrdu:
             return 59
        case .NotoSansKannada:
             return 60
        case .NotoSansMyanmar:
             return 61
        case .NotoSansOriya:
             return 62
        case .Optima:
             return 63
        case .Palatino:
             return 64
        case .Papyrus:
             return 65
        case .PartyLET:
             return 66
        case .PingFangHK:
             return 67
        case .PingFangSC:
             return 68
        case .PingFangTC:
             return 69
        case .Rockwell:
             return 70
        case .SavoyeLET:
             return 71
        case .SinhalaSangamMN:
             return 72
        case .SnellRoundhand:
             return 73
        case .Symbol:
             return 74
        case .TamilSangamMN:
             return 76
        case .Thonburi:
             return 77
        case .TimesNewRoman:
             return 78
        case .TrebuchetMS:
             return 79
        case .Verdana:
             return 80
        case .ZapfDingbats:
             return 81
        case .Zapfino:
             return 82
        }
    }
}
