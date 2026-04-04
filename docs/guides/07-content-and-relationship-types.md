---
title: "Content and Relationship Types"
group: "Guides"
category: "Reference"
---

# Content and Relationship Types

Every part in an Open XML package has a **content type** (a MIME type identifying what kind of content the part holds) and is connected to other parts through **relationships** (each identified by a relationship type URI). OpenXmlSdkTs provides two lookup objects that enumerate all standard types.

## ContentType

The `ContentType` object maps human-readable property names to their full MIME type strings.

### Type Aliases

- **`ContentTypeKey`** -- Union of all property names (e.g., `"mainDocument"`, `"worksheet"`, `"slide"`).
- **`ContentTypeValue`** -- Union of all MIME type string values.

### Usage

```typescript
import { ContentType } from "openxmlsdkts";

// Query parts by content type
const parts = await pkg.getPartsByContentType(ContentType.mainDocument);

// Use when adding a new part
await mainPart.addPart("/word/comments.xml", ContentType.wordprocessingComments, "xml", xDoc);
```

### All ContentType Properties

The following 87 properties are available on the `ContentType` object:

| Property | Property | Property |
|----------|----------|----------|
| `calculationChain` | `cellMetadata` | `chart` |
| `chartColorStyle` | `chartDrawing` | `chartsheet` |
| `chartStyle` | `commentAuthors` | `connections` |
| `coreFileProperties` | `customFileProperties` | `customization` |
| `customProperty` | `customXmlProperties` | `diagramColors` |
| `diagramData` | `diagramLayoutDefinition` | `diagramPersistLayout` |
| `diagramStyle` | `dialogsheet` | `digitalSignatureOrigin` |
| `documentSettings` | `drawings` | `endnotes` |
| `excelAttachedToolbars` | `extendedFileProperties` | `externalWorkbook` |
| `fontData` | `fontTable` | `footer` |
| `footnotes` | `gif` | `glossaryDocument` |
| `handoutMaster` | `header` | `jpeg` |
| `mainDocument` | `notesMaster` | `notesSlide` |
| `numberingDefinitions` | `pict` | `pivotTable` |
| `pivotTableCacheDefinition` | `pivotTableCacheRecords` | `png` |
| `presentation` | `presentationProperties` | `presentationTemplate` |
| `queryTable` | `relationships` | `ribbonAndBackstageCustomizations` |
| `sharedStringTable` | `singleCellTable` | `slicerCache` |
| `slicers` | `slide` | `slideComments` |
| `slideLayout` | `slideMaster` | `slideShow` |
| `slideSyncData` | `styles` | `tableDefinition` |
| `tableStyles` | `theme` | `themeOverride` |
| `tiff` | `trueTypeFont` | `userDefinedTags` |
| `viewProperties` | `vmlDrawing` | `volatileDependencies` |
| `webSettings` | `wordAttachedToolbars` | `wordprocessingComments` |
| `wordprocessingTemplate` | `workbook` | `workbookRevisionHeader` |
| `workbookRevisionLog` | `workbookStyles` | `workbookTemplate` |
| `workbookUserData` | `worksheet` | `worksheetComments` |
| `worksheetSortMap` | `xmlSignature` | |

## RelationshipType

The `RelationshipType` object maps human-readable property names to their full relationship type URIs.

### Type Aliases

- **`RelationshipTypeKey`** -- Union of all property names (e.g., `"styles"`, `"image"`, `"slide"`).
- **`RelationshipTypeValue`** -- Union of all relationship type URI values.

### Usage

```typescript
import { RelationshipType } from "openxmlsdkts";

// Get a part by relationship type
const stylePart = await mainPart.getPartByRelationshipType(RelationshipType.styles);

// Get all image relationships
const imageRels = await mainPart.getRelationshipsByRelationshipType(RelationshipType.image);

// Add a relationship
await mainPart.addRelationship("rId20", RelationshipType.image, "media/image1.png");
```

### All RelationshipType Properties

The following 80 properties are available on the `RelationshipType` object:

| Property | Property | Property |
|----------|----------|----------|
| `alternativeFormatImport` | `calculationChain` | `cellMetadata` |
| `chart` | `chartColorStyle` | `chartDrawing` |
| `chartsheet` | `chartStyle` | `commentAuthors` |
| `connections` | `coreFileProperties` | `customFileProperties` |
| `customization` | `customProperty` | `customXml` |
| `customXmlMappings` | `customXmlProperties` | `diagramColors` |
| `diagramData` | `diagramLayoutDefinition` | `diagramPersistLayout` |
| `diagramStyle` | `dialogsheet` | `digitalSignatureOrigin` |
| `documentSettings` | `drawings` | `endnotes` |
| `excelAttachedToolbars` | `extendedFileProperties` | `externalWorkbook` |
| `font` | `fontTable` | `footer` |
| `footnotes` | `glossaryDocument` | `handoutMaster` |
| `header` | `image` | `mainDocument` |
| `notesMaster` | `notesSlide` | `numberingDefinitions` |
| `pivotTable` | `pivotTableCacheDefinition` | `pivotTableCacheRecords` |
| `presentation` | `presentationProperties` | `queryTable` |
| `ribbonAndBackstageCustomizations` | `sharedStringTable` | `singleCellTable` |
| `slide` | `slideComments` | `slideLayout` |
| `slideMaster` | `slideSyncData` | `styles` |
| `tableDefinition` | `tableStyles` | `timeLine` |
| `theme` | `themeOverride` | `thumbnail` |
| `userDefinedTags` | `viewProperties` | `vmlDrawing` |
| `volatileDependencies` | `webSettings` | `wordAttachedToolbars` |
| `wordprocessingComments` | `workbook` | `workbookRevisionHeader` |
| `workbookRevisionLog` | `workbookStyles` | `workbookUserData` |
| `worksheet` | `worksheetComments` | `worksheetSortMap` |
| `xmlSignature` | | |

## Querying with Types

Both objects are commonly used with the generic navigation methods on packages and parts:

```typescript
import { ContentType, RelationshipType } from "openxmlsdkts";

// Find all worksheet parts by content type
const worksheets = await workbookPart.getPartsByContentType(ContentType.worksheet);

// Find all image parts by relationship type
const images = await mainPart.getPartsByRelationshipType(RelationshipType.image);

// Get the single styles part by relationship type
const stylesPart = await mainPart.getPartByRelationshipType(RelationshipType.styles);
```
