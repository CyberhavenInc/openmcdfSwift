import Common
import Foundation
import Utility

public enum VTPropertyType: UInt16 {
    case VT_EMPTY = 0
    case VT_NULL = 1
    case VT_I2 = 2
    case VT_I4 = 3
    case VT_R4 = 4
    case VT_R8 = 5
    case VT_CY = 6
    case VT_DATE = 7
    case VT_BSTR = 8
    case VT_ERROR = 10
    case VT_BOOL = 11
    case VT_VARIANT_VECTOR = 12
    case VT_DECIMAL = 14
    case VT_I1 = 16
    case VT_UI1 = 17
    case VT_UI2 = 18
    case VT_UI4 = 19
    case VT_I8 = 20
    case VT_UI8 = 21
    case VT_INT = 22
    case VT_UINT = 23
    case VT_LPSTR = 30
    case VT_LPWSTR = 31
    case VT_FILETIME = 64
    case VT_BLOB = 65
    case VT_STREAM = 66
    case VT_STORAGE = 67
    case VT_STREAMED_OBJECT = 68
    case VT_STORED_OBJECT = 69
    case VT_BLOB_OBJECT = 70
    case VT_CF = 71
    case VT_CLSID = 72
    case VT_VERSIONED_STREAM = 73
    case VT_VECTOR_HEADER = 4096
    case VT_ARRAY_HEADER = 8192
    case VT_VARIANT_ARRAY = 8204
}

public class PlainProperty<T: DataConvertible>: OleTypedProperty<T> {
    override func readScalarValue(stream: DataReader) -> T? {
        let data = stream.readData(ofLength: MemoryLayout<T>.size)
        return T(data: data)
    }
}

public final class VT_I1_Property: PlainProperty<Int8> {}
public final class VT_UI1_Property: PlainProperty<UInt8> {}
public final class VT_I2_Property: PlainProperty<Int16> {}
public final class VT_UI2_Property: PlainProperty<UInt16> {}
public final class VT_I4_Property: PlainProperty<Int32> {}
public final class VT_UI4_Property: PlainProperty<UInt32> {}
public final class VT_R4_Property: PlainProperty<Float> {}
public final class VT_R8_Property: PlainProperty<Double> {}

public class DateProperty: OleTypedProperty<Date> {
    override public var description: String {
        let dateformat = DateFormatter()
        dateformat.timeZone = TimeZone(abbreviation: "UTC")
        dateformat.dateFormat = "yyyy-MM-dd'T'HH:mm:ss'Z'"

        switch dimensions {
        case .scalar:
            if let value {
                return "\(dateformat.string(from: value))"
            }
        case .vector:
            if let values {
                var result = [String]()
                for value in values {
                    result.append("\(dateformat.string(from: value))")
                }
                return "[\(result.joined(separator: ", "))]"
            }
        default:
            L.warning("Unsupported property dimension: \(dimensions)")
        }

        return "<nil>"
    }
}

public final class VT_EMPTY_Property: OleTypedProperty<Void> {
    override func readScalarValue(stream: DataReader) {
        ()
    }
}

public final class VT_CY_Property: OleTypedProperty<Int64> {
    override func readScalarValue(stream: DataReader) -> Int64? {
        let data = stream.readData(ofLength: MemoryLayout<Int64>.size)
        if let value = Int64(data: data) {
            return value / 10000
        }

        return nil
    }
}

public final class VT_BOOL_Property: PlainProperty<UInt16> {
    override public var description: String {
        switch dimensions {
        case .scalar:
            if let value {
                return "\(value != 0 ? "true" : "false")"
            }
        case .vector:
            if let values {
                var result = [String]()
                for value in values {
                    result.append("\(value != 0 ? "true" : "false")")
                }
                return "[\(result.joined(separator: ", "))]"
            }
        default:
            L.warning("Unsupported property dimension: \(dimensions)")
        }

        return "<nil>"
    }
}

public class VT_LPSTR_Property: OleTypedProperty<String> {
    public let encoding: String.Encoding

    public init(vtType: UInt16, codePage: UInt16, isVariant: Bool) {
        encoding = codePageToEncoding(codePage)
        super.init(vtType: vtType, isVariant: isVariant)
    }

    override func readScalarValue(stream: DataReader) -> String? {
        let length: UInt32 = stream.read()
        let data = stream.readData(ofLength: Int(length))
        return String(data: data, encoding: encoding)?.trimmingCharacters(in: CharacterSet(charactersIn: "\0"))
    }
}

public final class VT_LPWSTR_Property: VT_LPSTR_Property {
    override func readScalarValue(stream: DataReader) -> String? {
        let length: UInt32 = stream.read()
        let data = stream.readData(ofLength: 2 * Int(length))
        return String(data: data, encoding: encoding)?.trimmingCharacters(in: CharacterSet(charactersIn: "\0"))
    }
}

public final class VT_FILETIME_Property: DateProperty {
    override func readScalarValue(stream: DataReader) -> Date? {
        let data = stream.readData(ofLength: MemoryLayout<Int64>.size)
        if let filetime = Int64(data: data) {
            let seconds = filetime / 10000000
            let epoch = seconds - 11644473600
            let date = Date(timeIntervalSince1970: TimeInterval(epoch))
            return date
        }

        return nil
    }
}

public final class VT_DATE_Property: DateProperty {
    override func readScalarValue(stream: DataReader) -> Date? {
        let data = stream.readData(ofLength: MemoryLayout<Double>.size)
        if let oaDate = Double(data: data) {
            // Number of days between 1 January 1900 and 1 January 1970.
            let dateDiff = 25569.0
            let epoch = (oaDate - dateDiff) * 24 * 3600
            let date = Date(timeIntervalSince1970: TimeInterval(epoch))
            return date
        }

        return nil
    }
}

public final class VT_VARIANT_VECTOR_Property: OleTypedProperty<OleProperty> {
    private let mCodePage: UInt16

    public init(vtType: UInt16, codePage: UInt16, isVariant: Bool) {
        mCodePage = codePage
        super.init(vtType: vtType, isVariant: isVariant)
    }

    override func readScalarValue(stream: DataReader) -> OleProperty? {
        let vType: UInt16 = stream.read()
        let _: UInt16 = stream.read()
        if let property = OlePropertyFactory.createProperty(vtType: vType, codePage: mCodePage, isVariant: true) {
            property.read(stream: stream)
            return property
        }

        return nil
    }
}
