// Copyright 2019-2020 CoreOffice contributors
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
//  Created by Max Desiatov on 27/10/2018.
//

/** An array of `Relationship` values. This type directly maps the internal XML structure of the
 `.xlsx` format.
 */
public struct Relationships: Codable, Equatable {
  public let items: [Relationship]

  enum CodingKeys: String, CodingKey {
    case items = "relationship"
  }
}

/** Relationship to an entity stored in a given `.xlsx` archive. These can be worksheets,
 chartsheets, thumbnails and a few other internal entities. Most of the time users of CoreXLSX
 wouldn't need to handle relationships directly.
 */
public struct Relationship: Codable, Equatable {
  public enum SchemaType: Codable, Equatable {
    case calcChain
    case officeDocument
    case extendedProperties
    case packageCoreProperties
    case coreProperties
    case connections
    case worksheet
    case chartsheet
    case sharedStrings
    case styles
    case theme
    case pivotCache
    case metadataThumbnail
    case customProperties
    case externalLink
    case customXml
    case person
    case webExtensionTaskPanes
    case googleWorkbookMetadata
    case purlOCLC
    case classificationLabels
    case richDataStructure
    case richDataTypes
    case richDataWebImage
    // Forward compatibility: handle unknown schema types
    case unknown(String)
    
    private static let knownSchemas: [String: SchemaType] = [
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain": .calcChain,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument": .officeDocument,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties": .extendedProperties,
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties": .packageCoreProperties,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties": .coreProperties,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections": .connections,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet": .worksheet,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet": .chartsheet,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings": .sharedStrings,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles": .styles,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme": .theme,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition": .pivotCache,
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail": .metadataThumbnail,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties": .customProperties,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink": .externalLink,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml": .customXml,
      "http://schemas.microsoft.com/office/2017/10/relationships/person": .person,
      "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes": .webExtensionTaskPanes,
      "http://customschemas.google.com/relationships/workbookmetadata": .googleWorkbookMetadata,
      "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties": .purlOCLC,
      "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels": .classificationLabels,
      "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure": .richDataStructure,
      "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue": .richDataTypes,
      "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueWebImage": .richDataWebImage
    ]
    
    public init(from decoder: Decoder) throws {
      let container = try decoder.singleValueContainer()
      let stringValue = try container.decode(String.self)
      
      self = Self.knownSchemas[stringValue] ?? .unknown(stringValue)
    }
    
    public func encode(to encoder: Encoder) throws {
      var container = encoder.singleValueContainer()
      
      switch self {
      case .unknown(let value):
        try container.encode(value)
      default:
        // Find the string value for known cases
        for (key, schemaType) in Self.knownSchemas {
          if schemaType == self {
            try container.encode(key)
            return
          }
        }
        throw EncodingError.invalidValue(self, EncodingError.Context(codingPath: encoder.codingPath, debugDescription: "Unknown schema type"))
      }
    }
    
    /// Returns the raw string value for this schema type
    public var rawValue: String {
      switch self {
      case .unknown(let value):
        return value
      default:
        for (key, schemaType) in Self.knownSchemas {
          if schemaType == self {
            return key
          }
        }
        return ""
      }
    }
    
    /// Returns true if this is a known schema type, false if unknown
    public var isKnown: Bool {
      switch self {
      case .unknown:
        return false
      default:
        return true
      }
    }
    
    /// Returns the unknown schema string if this is an unknown type, nil otherwise
    public var unknownSchema: String? {
      switch self {
      case .unknown(let value):
        return value
      default:
        return nil
      }
    }
    
    /// Check if this schema type is a worksheet-related type
    public var isWorksheetRelated: Bool {
      switch self {
      case .worksheet, .chartsheet:
        return true
      default:
        return false
      }
    }
    
    /// Check if this schema type is a Microsoft Office 365/2020+ feature
    public var isModernOfficeFeature: Bool {
      switch self {
      case .classificationLabels, .richDataStructure, .richDataTypes, .richDataWebImage, .person:
        return true
      case .unknown(let value):
        return value.contains("schemas.microsoft.com/office/20")
      default:
        return false
      }
    }
  }

  /// The identifier for this entity.
  public let id: String

  /// The type of this entity.
  public let type: SchemaType

  /// The path to this entity in the `.xlsx` archive.
  public let target: String

  func path(from root: String) -> String {
    Path(target).isRoot ? target : "\(root)/\(target)"
  }
}
