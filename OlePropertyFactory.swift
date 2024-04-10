import Common
import Foundation

enum OlePropertyFactory {
    public static func createProperty(vtType: UInt16, codePage: UInt16, isVariant: Bool = false) -> OleProperty? {
        guard let propType = VTPropertyType(rawValue: vtType & 0xff) else {
            L.error("Unknown property type: \(vtType & 0xff)")
            return nil
        }

        // MS Office GUI applications can set only I4, Bool, LPSTR, and FILETIME properties.
        switch propType {
        case .VT_EMPTY, .VT_NULL:
            return VT_EMPTY_Property(vtType: vtType, isVariant: isVariant)
        case .VT_I1:
            return VT_I1_Property(vtType: vtType, isVariant: isVariant)
        case .VT_UI1:
            return VT_UI1_Property(vtType: vtType, isVariant: isVariant)
        case .VT_I2:
            return VT_I2_Property(vtType: vtType, isVariant: isVariant)
        case .VT_UI2:
            return VT_UI2_Property(vtType: vtType, isVariant: isVariant)
        case .VT_I4, .VT_INT:
            return VT_I4_Property(vtType: vtType, isVariant: isVariant)
        case .VT_UI4, .VT_UINT:
            return VT_UI4_Property(vtType: vtType, isVariant: isVariant)
        case .VT_R4:
            return VT_R4_Property(vtType: vtType, isVariant: isVariant)
        case .VT_R8:
            return VT_R8_Property(vtType: vtType, isVariant: isVariant)
        case .VT_CY:
            return VT_CY_Property(vtType: vtType, isVariant: isVariant)
        case .VT_BOOL:
            return VT_BOOL_Property(vtType: vtType, isVariant: isVariant)
        case .VT_LPSTR, .VT_BSTR:
            return VT_LPSTR_Property(vtType: vtType, codePage: codePage, isVariant: isVariant)
        case .VT_LPWSTR:
            return VT_LPWSTR_Property(vtType: vtType, codePage: codePage, isVariant: isVariant)
        case .VT_FILETIME:
            return VT_FILETIME_Property(vtType: vtType, isVariant: isVariant)
        case .VT_DATE:
            return VT_DATE_Property(vtType: vtType, isVariant: isVariant)
        case .VT_VARIANT_VECTOR:
            return VT_VARIANT_VECTOR_Property(vtType: vtType, codePage: codePage, isVariant: isVariant)
        default:
            L.error("Unsupported property: \(propType)")
            return nil
        }
    }
}
