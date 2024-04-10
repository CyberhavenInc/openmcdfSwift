import Common
import Foundation
import Utility

// Ported .NET OpenMcdf OLEProperties implementation:
// https://github.com/ironfede/openmcdf/tree/master/sources/OpenMcdf.Extensions/OLEProperties

public enum PropertyType {
    case typedPropertyValue
    case dictionaryProperty
}

public enum PropertyDimensions {
    case scalar
    case vector
    case array
}

public protocol OleProperty: CustomStringConvertible {
    var id: UInt32 { get }
    var name: String? { get }

    func setId(_ id: UInt32)
    func setName(_ name: String?)
    func read(stream: DataReader)
}

public class OleTypedProperty<T>: OleProperty {
    public var id: UInt32
    public var name: String?
    public let isVariant: Bool
    public let propertyType: PropertyType = .typedPropertyValue
    public let vtType: UInt16
    public let dimensions: PropertyDimensions
    private(set) var value: T?
    private(set) var values: [T]?

    public init(vtType: UInt16, isVariant: Bool) {
        id = 0
        self.isVariant = isVariant
        self.vtType = vtType
        if vtType & VTPropertyType.VT_VECTOR_HEADER.rawValue != VTPropertyType.VT_EMPTY.rawValue {
            dimensions = .vector
        } else if vtType & VTPropertyType.VT_ARRAY_HEADER.rawValue != VTPropertyType.VT_EMPTY.rawValue {
            dimensions = .array
        } else {
            dimensions = .scalar
        }
    }

    public func setId(_ id: UInt32) {
        self.id = id
    }

    public func setName(_ name: String?) {
        self.name = name
    }

    public var description: String {
        switch dimensions {
        case .scalar:
            if let value {
                if let container = value as? OleProperty {
                    return container.description
                }
                return "\(value)"
            }
        case .vector:
            if let values {
                var result = [String]()
                for value in values {
                    if let container = value as? OleProperty {
                        result.append(container.description)
                        continue
                    }
                    result.append("\(value)")
                }
                return "[\(result.joined(separator: ", "))]"
            }
        default:
            L.error("Unsupported property dimension: \(dimensions)")
        }

        return "<nil>"
    }

    public func read(stream: DataReader) {
        let position = stream.position
        switch dimensions {
        case .scalar:
            value = readScalarValue(stream: stream)
        case .vector:
            values = []
            let count: UInt32 = stream.read()
            for _ in 0 ..< count {
                if let data = readScalarValue(stream: stream) {
                    values!.append(data)
                }
            }
        default:
            L.error("Unsupported property dimension: \(dimensions)")
            return
        }

        let padding = (stream.position - position) % 4
        if padding != 0, !isVariant, stream.position + padding < stream.totalBytes {
            stream.seek(toOffset: stream.position + padding)
        }
    }

    func readScalarValue(stream: DataReader) -> T? {
        fatalError("Subclasses must implement 'readScalarValue'")
    }
}

public func codePageToEncoding(_ codePage: UInt16) -> String.Encoding {
    switch codePage {
    case 1200:
        return .utf16
    case 1250:
        return .windowsCP1250
    case 1251:
        return .windowsCP1251
    case 1252:
        return .windowsCP1252
    case 1253:
        return .windowsCP1253
    case 1254:
        return .windowsCP1254
    default:
        L.warning("Unsupported encoding \(codePage), using default utf8")
        return .utf8
    }
}

private final class EntryWithOffset {
    public let id: UInt32
    public let offset: Int

    public init(id: UInt32, offset: UInt32) {
        self.id = id
        self.offset = Int(offset)
    }
}

public final class OlePropertySet {
    private let mSize: UInt32
    private let mNumProperties: UInt32
    private var mCodePage: UInt16 = 0
    private var mEntries = [EntryWithOffset]()
    private var mIdToNameMap = [UInt32: String]()

    public let id: UUID
    public var properties = [OleProperty]()
    public var hasNamedProps: Bool {
        !mIdToNameMap.isEmpty
    }

    public init(stream: DataReader, baseOffset: Int, fmtid: Data) {
        if stream.position != baseOffset {
            stream.seek(toOffset: baseOffset)
        }

        id = UUID(fromGUIDBytes: fmtid) ?? UUID()
        mSize = stream.read()
        mNumProperties = stream.read()

        for _ in 0 ..< mNumProperties {
            let id: UInt32 = stream.read()
            let offset: UInt32 = stream.read()
            mEntries.append(EntryWithOffset(id: id, offset: offset))
        }

        loadContext(stream: stream, offset: baseOffset)
        loadNameDictionary(stream: stream, baseOffset: baseOffset)
        readProperties(stream: stream, baseOffset: baseOffset)
    }

    private func loadContext(stream: DataReader, offset: Int) {
        let position = stream.position
        defer {
            stream.seek(toOffset: position)
        }

        if let firstEntry = mEntries.first(where: { $0.id == 1 }) {
            stream.seek(toOffset: offset + firstEntry.offset)
            let _: UInt16 = stream.read()
            let _: UInt16 = stream.read()
            mCodePage = stream.read()
        }
    }

    private func loadNameDictionary(stream: DataReader, baseOffset: Int) {
        guard let dictionaryEntry = mEntries.first(where: { $0.id == 0 }) else {
            return
        }

        let encoding = codePageToEncoding(mCodePage)

        stream.seek(toOffset: baseOffset + dictionaryEntry.offset)
        let position = stream.position

        let count: UInt32 = stream.read()
        for _ in 0 ..< count {
            let propId: UInt32 = stream.read()
            let length: UInt32 = stream.read()

            var data: Data
            if encoding != .utf16 {
                data = stream.readData(ofLength: Int(length))
            } else {
                data = stream.readData(ofLength: 2 * Int(length))
                let padding = (stream.position - position) % 4
                if padding > 0 {
                    stream.seek(toOffset: stream.position + padding)
                }
            }

            let name = String(data: data, encoding: encoding)
            if let name = name?.trimmingCharacters(in: CharacterSet(charactersIn: "\0")) {
                mIdToNameMap[propId] = name
            }
        }
    }

    private func readProperties(stream: DataReader, baseOffset: Int) {
        for entry in mEntries {
            stream.seek(toOffset: baseOffset + entry.offset)
            if entry.id != 0 {
                let propTypeRaw: UInt16 = stream.read()
                let _: UInt16 = stream.read()

                guard let property = OlePropertyFactory.createProperty(vtType: propTypeRaw, codePage: mCodePage) else {
                    continue
                }

                property.setId(entry.id)
                property.setName(mIdToNameMap[entry.id])
                property.read(stream: stream)

                properties.append(property)
            }
        }
    }
}

public final class OlePropertyCollection {
    // https://docs.microsoft.com/en-us/windows/win32/stg/predefined-property-set-format-identifiers
    public static let fmtidSummaryInformation = UUID(uuidString: "F29F85E0-4FF9-1068-AB91-08002B27B3D9")!
    public static let fmtidDocSummaryInformation = UUID(uuidString: "D5CDD502-2E9C-101B-9397-08002B2CF9AE")!
    public static let fmtidUserDefinedProperties = UUID(uuidString: "D5CDD505-2E9C-101B-9397-08002B2CF9AE")!

    public let byteOrder: UInt16
    public let version: UInt16
    public let systemIdentifier: UInt32
    public let clsid: Data
    public let numPropertySets: UInt32
    private(set) var propSetDescriptors = [(fmtid: Data, offset: UInt32)]()
    private(set) var propertySets = [OlePropertySet]()

    public init(stream: DataReader) {
        byteOrder = stream.read()
        version = stream.read()
        systemIdentifier = stream.read()
        clsid = stream.readData(ofLength: 16)
        numPropertySets = stream.read()

        var fmtid = stream.readData(ofLength: 16)
        var offset: UInt32 = stream.read()
        propSetDescriptors.append((fmtid: fmtid, offset: offset))

        if numPropertySets == 2 {
            fmtid = stream.readData(ofLength: 16)
            offset = stream.read()
            propSetDescriptors.append((fmtid: fmtid, offset: offset))
        }

        for desc in propSetDescriptors {
            let propSet = OlePropertySet(stream: stream, baseOffset: Int(desc.offset), fmtid: desc.fmtid)
            propertySets.append(propSet)
        }
    }

    public func getAllProperties() -> [OleProperty] {
        var result = [OleProperty]()
        for propSet in propertySets {
            result.append(contentsOf: propSet.properties)
        }
        return result
    }

    public func getCustomProperties() -> [OleProperty] {
        var result = [OleProperty]()
        for propSet in propertySets where propSet.id == OlePropertyCollection.fmtidUserDefinedProperties {
            result.append(contentsOf: propSet.properties)
        }
        return result
    }
}
