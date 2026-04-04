# OpenXmlSdkTs — Open XML SDK for TypeScript

A TypeScript library for reading, writing, and manipulating Office Open XML documents (`.docx`, `.xlsx`, `.pptx`) in Node.js and browser environments. Inspired by the .NET [Open-Xml-Sdk](https://github.com/dotnet/Open-Xml-Sdk), this library brings the same familiar programming model to the TypeScript ecosystem.

## Documentation

Full API reference and guides are published at **[ericwhitedev.github.io/OpenXmlSdkTs](https://ericwhitedev.github.io/OpenXmlSdkTs/)**.

To build the documentation locally:

```bash
npm run docs
```

Output is written to `docs/api/`. Open `docs/api/index.html` in a browser.

## Why OpenXmlSdkTs?

- **Full document format support** — Work with Word, Excel, and PowerPoint files at the XML level.
- **Three I/O modes** — Open and save documents as binary blobs (via JSZip), Flat OPC XML strings (the required format when building Office JavaScript/TypeScript add-ins), or Base64 strings.
- **Friendly content & relationship types** — All Open XML content types and relationship types are referenced using readable labels (e.g. `RelationshipType.styles`, `ContentType.mainDocument`) instead of long, error-prone URIs.
- **Pre-initialized namespace, element, and attribute names** — Static classes (`W`, `S`, `P`, `A`, etc.) provide pre-initialized `XName` and `XNamespace` objects for every element and attribute in the Open XML specification. Because `XNamespace` and `XName` objects are *atomized* (two objects with the same namespace and local name are the same object), this gives excellent performance when querying and modifying markup.
- **Built on LINQ to XML for TypeScript** — Powered by [`ltxmlts`](https://www.npmjs.com/package/ltxmlts), a faithful TypeScript port of .NET's LINQ to XML. Query and transform XML with `elements()`, `descendants()`, `attributes()`, and the rest of the LINQ to XML API you already know.
- **Intuitive part navigation** — Navigate from the package to the main document part, then from part to part using typed methods that mirror the .NET SDK: `mainDocumentPart()`, `styleDefinitionsPart()`, `worksheetParts()`, `slideParts()`, and many more.
- **Lightweight** — Only two runtime dependencies: `jszip` and `ltxmlts`.
- **MIT licensed** — Free for commercial and open-source use.  This is the same license as the C#/dotnet Open-Xml-Sdk.

## Installation

```bash
npm install openxmlsdkts
```

Both `openxmlsdkts` and its companion library [`ltxmlts`](https://www.npmjs.com/package/ltxmlts) are available on [npmjs](https://www.npmjs.com/package/openxmlsdkts).

## Class Hierarchy

```
OpenXmlPackage                      Base class for all Office document packages
├── WmlPackage                      Word (.docx) packages
├── SmlPackage                      Excel (.xlsx) packages
└── PmlPackage                      PowerPoint (.pptx) packages

OpenXmlPart                         Base class for all parts within a package
├── WmlPart                         Word parts (document, styles, headers, footers, etc.)
├── SmlPart                         Excel parts (workbook, worksheets, charts, etc.)
└── PmlPart                         PowerPoint parts (presentation, slides, masters, etc.)

OpenXmlRelationship                 Represents a relationship between a package/part and a target

ContentType                         Static lookup of content type labels → URIs
RelationshipType                    Static lookup of relationship type labels → URIs

Namespace Classes (static)          Pre-initialized XNamespace / XName objects
├── W                               WordprocessingML
├── S                               SpreadsheetML
├── P                               PresentationML
├── A                               DrawingML
├── R, M, MC, W14, ...             30+ additional namespace classes
└── NoNamespace                     Attributes with no namespace
```

## Quick Start

### Load a Word Document, Modify It, and Save

```typescript
import * as fs from "fs";
import { WmlPackage, W, WmlPart } from "openxmlsdkts";

async function capitalizeFirstWord() {
  // 1. Load the document from a file
  const buffer = fs.readFileSync("my-document.docx");
  const blob = new Blob([buffer]);
  const doc = await WmlPackage.open(blob);

  // 2. Get the main document part and its XML
  const mainPart: WmlPart | undefined = await doc.mainDocumentPart();
  if (!mainPart) throw new Error("No main document part found");

  const xDoc = await mainPart.getXDocument();

  // 3. Find the first word of the first run of text and capitalize it
  const body = xDoc.root!.element(W.body);
  if (body) {
    const firstParagraph = body.element(W.p);
    if (firstParagraph) {
      const firstRun = firstParagraph.element(W.r);
      if (firstRun) {
        const textElement = firstRun.element(W.t);
        if (textElement && textElement.value) {
          const words = textElement.value.split(" ");
          words[0] = words[0].toUpperCase();
          textElement.value = words.join(" ");
        }
      }
    }
  }

  // 4. Save the modified XML back to the part
  mainPart.putXDocument(xDoc);

  // 5. Save the document back to a file
  const savedBlob = await doc.saveToBlobAsync();
  const arrayBuffer = await savedBlob.arrayBuffer();
  fs.writeFileSync("my-document-modified.docx", Buffer.from(arrayBuffer));
}

capitalizeFirstWord();
```

## Opening and Saving Documents

OpenXmlSdkTs supports three document formats. The `open()` method auto-detects the format from the input.

### Binary (Blob)

The standard `.docx` / `.xlsx` / `.pptx` ZIP-based format.

```typescript
// Open
const buffer = fs.readFileSync("report.docx");
const doc = await WmlPackage.open(new Blob([buffer]));

// Save
const blob = await doc.saveToBlobAsync();
```

### Flat OPC (XML String)

A single-file XML representation of the entire package. This is the format you must work with when building **Office JavaScript/TypeScript add-in** applications.

```typescript
// Open
const flatOpcXml = fs.readFileSync("report.xml", "utf-8");
const doc = await WmlPackage.open(flatOpcXml);

// Save
const flatOpc: string = await doc.saveToFlatOpcAsync();
```

### Base64 String

Convenient for APIs, data URIs, and serialization scenarios.

```typescript
// Open
const base64 = fs.readFileSync("report.txt", "utf-8"); // base64-encoded docx
const doc = await WmlPackage.open(base64);

// Save
const savedBase64: string = await doc.saveToBase64Async();
```

## Navigating Parts

Each package type provides typed methods to access its constituent parts.

### Word Documents (WmlPackage / WmlPart)

```typescript
const doc = await WmlPackage.open(blob);

// Package-level
const mainPart = await doc.mainDocumentPart();
const contentParts = await doc.contentParts(); // main + headers + footers + footnotes + endnotes

// Part-level navigation
const styles = await mainPart!.styleDefinitionsPart();
const theme = await mainPart!.themePart();
const numbering = await mainPart!.numberingDefinitionsPart();
const headers = await mainPart!.headerParts();
const footers = await mainPart!.footerParts();
const comments = await mainPart!.wordprocessingCommentsPart();
const fonts = await mainPart!.fontTablePart();
const settings = await mainPart!.documentSettingsPart();
```

### Excel Spreadsheets (SmlPackage / SmlPart)

```typescript
const doc = await SmlPackage.open(blob);
const workbook = await doc.workbookPart();

const worksheets = await workbook!.worksheetParts();
const sharedStrings = await workbook!.sharedStringTablePart();
const wbStyles = await workbook!.workbookStylesPart();
const charts = await workbook!.chartsheetParts();
const pivotTables = await workbook!.pivotTableParts();
```

### PowerPoint Presentations (PmlPackage / PmlPart)

```typescript
const doc = await PmlPackage.open(blob);
const presentation = await doc.presentationPart();

const slides = await presentation!.slideParts();
const masters = await presentation!.slideMasterParts();
const layout = await presentation!.slideLayoutPart();
const notesMaster = await presentation!.notesMasterPart();
```

## Working with XML

All XML manipulation uses the LINQ to XML API from [`ltxmlts`](https://www.npmjs.com/package/ltxmlts). The pre-initialized namespace classes make queries concise and performant.

```typescript
import { WmlPackage, W, NoNamespace, XElement } from "openxmlsdkts";

const doc = await WmlPackage.open(blob);
const mainPart = await doc.mainDocumentPart();
const xDoc = await mainPart!.getXDocument();

// Query all paragraphs
const paragraphs = xDoc.root!.element(W.body)!.elements(W.p);

// Find paragraphs with a specific style
for (const para of paragraphs) {
  const pPr = para.element(W.pPr);
  const pStyle = pPr?.element(W.pStyle);
  const styleVal = pStyle?.attribute(NoNamespace.val)?.value;
  if (styleVal === "Heading1") {
    console.log("Found Heading1 paragraph");
  }
}

// Add a new paragraph
const newPara = new XElement(W.p,
  new XElement(W.r,
    new XElement(W.t, "Hello from OpenXmlSdkTs!")
  )
);
xDoc.root!.element(W.body)!.add(newPara);

// Write the modified XML back to the part
mainPart!.putXDocument(xDoc);
```

## Content Types and Relationship Types

Instead of using long URIs, use the `ContentType` and `RelationshipType` static lookups.

```typescript
import { ContentType, RelationshipType } from "openxmlsdkts";

// Navigate parts by relationship type
const stylePart = await mainPart!.getPartByRelationshipType(RelationshipType.styles);
const themePart = await mainPart!.getPartByRelationshipType(RelationshipType.theme);
const imageParts = await mainPart!.getPartsByRelationshipType(RelationshipType.image);

// Query relationships
const rels = await mainPart!.getRelationships();
for (const rel of rels) {
  console.log(rel.getType());    // e.g. the styles relationship URI
  console.log(rel.getTarget());  // e.g. "styles.xml"
}
```

## Pre-Initialized Namespace Classes

Over 40 static classes cover every namespace in the Open XML specification:

| Class | Namespace | Description |
|-------|-----------|-------------|
| `W` | `wordprocessingml/2006/main` | WordprocessingML elements and attributes |
| `S` | `spreadsheetml/2006/main` | SpreadsheetML elements and attributes |
| `P` | `presentationml/2006/main` | PresentationML elements and attributes |
| `A` | `drawingml/2006/main` | DrawingML elements and attributes |
| `R` | `officeDocument/2006/relationships` | Relationship attributes |
| `M` | `officeDocument/2006/math` | Office Math |
| `W14` | `microsoft.com/office/word/2010/wordml` | Word 2010 extensions |
| `MC` | `markup-compatibility/2006` | Markup Compatibility |
| `NoNamespace` | *(none)* | Attributes with no namespace (e.g. `val`, `id`) |
| ... | ... | 30+ additional namespace classes |

Each class exposes a `namespace` property and `XName` properties for every element and attribute:

```typescript
W.namespace  // XNamespace for WordprocessingML
W.body       // XName for <w:body>
W.p          // XName for <w:p>
W.r          // XName for <w:r>
W.t          // XName for <w:t>
```

Because `XName` and `XNamespace` objects are atomized, equality checks are identity checks (`===`), giving excellent query performance.

## Use with Office Add-ins

Flat OPC is the document format you must use when building **Word, Excel, and PowerPoint JavaScript/TypeScript add-ins**:

```typescript
// In an Office Add-in — get the document as Flat OPC via the Office.js API
// then manipulate with OpenXmlSdkTs
const doc = await WmlPackage.open(flatOpcString);
const mainPart = await doc.mainDocumentPart();
// ... modify the document ...
const modifiedFlatOpc = await doc.saveToFlatOpcAsync();
// Set the modified Flat OPC back via the Office.js API
```

## API Reference

### Packages

| Class | Description | Key Methods |
|-------|-------------|-------------|
| `WmlPackage` | Word documents | `open()`, `mainDocumentPart()`, `contentParts()`, `saveToBase64Async()`, `saveToBlobAsync()`, `saveToFlatOpcAsync()` |
| `SmlPackage` | Excel spreadsheets | `open()`, `workbookPart()`, `saveToBase64Async()`, `saveToBlobAsync()`, `saveToFlatOpcAsync()` |
| `PmlPackage` | PowerPoint presentations | `open()`, `presentationPart()`, `saveToBase64Async()`, `saveToBlobAsync()`, `saveToFlatOpcAsync()` |

### Parts

| Class | Notable Navigation Methods |
|-------|---------------------------|
| `WmlPart` | `headerParts()`, `footerParts()`, `styleDefinitionsPart()`, `themePart()`, `numberingDefinitionsPart()`, `fontTablePart()`, `documentSettingsPart()`, `endnotesPart()`, `footnotesPart()`, `wordprocessingCommentsPart()`, `glossaryDocumentPart()` |
| `SmlPart` | `worksheetParts()`, `chartsheetParts()`, `sharedStringTablePart()`, `workbookStylesPart()`, `calculationChainPart()`, `pivotTableParts()`, `tableDefinitionParts()` |
| `PmlPart` | `slideParts()`, `slideMasterParts()`, `slideLayoutPart()`, `notesMasterPart()`, `notesSlidePart()`, `handoutMasterPart()`, `commentAuthorsPart()`, `presentationPropertiesPart()`, `viewPropertiesPart()` |

### Common Methods (All Parts)

| Method | Description |
|--------|-------------|
| `getXDocument()` | Get the part's XML as an `XDocument` |
| `putXDocument(xDoc)` | Write modified XML back to the part |
| `getRelationships()` | Get all relationships from this part |
| `getPartByRelationshipType(type)` | Get a single related part by relationship type |
| `getPartsByRelationshipType(type)` | Get all related parts by relationship type |
| `getUri()` | Get the part's URI within the package |
| `getContentType()` | Get the part's content type |

## Dependencies

| Package | Purpose |
|---------|---------|
| [`ltxmlts`](https://www.npmjs.com/package/ltxmlts) | LINQ to XML for TypeScript — XML querying and manipulation |
| [`jszip`](https://www.npmjs.com/package/jszip) | ZIP compression for binary Open XML packages |

## Requirements

- **Node.js** 18+ or modern browser with `Blob` support
- **TypeScript** 5.0+ (recommended)

## License

MIT License. Copyright (c) 2026 Eric White.

## Author

**Eric White**
- [ericwhite.com](https://www.ericwhite.com)
- [LinkedIn](https://linkedin.com/in/ericwhitedev)
- [eric@ericwhite.com](mailto:eric@ericwhite.com)
