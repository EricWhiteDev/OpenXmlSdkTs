---
title: "Namespace Classes"
group: "Guides"
category: "Reference"
---

# Namespace Classes

OpenXmlSdkTs provides over 48 namespace classes, each containing pre-atomized `XName` constants for every element and attribute defined in that namespace of the Open XML specification.

## What Are Atomized Names?

In XML processing, element and attribute names are typically compared as strings. OpenXmlSdkTs pre-creates all names as `XName` objects at class load time. Because `XName` instances are interned (the same namespace + local name always returns the same object), comparisons become O(1) identity checks instead of O(n) string comparisons.

This also eliminates typos -- if you misspell a name, you get a compile-time error instead of a silent query that matches nothing.

## Pattern

Every namespace class follows the same pattern:

```typescript
export class W {
  static readonly namespace: XNamespace = XNamespace.get(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  );

  static readonly body: XName = W.namespace.getName("body");
  static readonly p: XName = W.namespace.getName("p");
  static readonly r: XName = W.namespace.getName("r");
  static readonly t: XName = W.namespace.getName("t");
  // ... hundreds more
}
```

Each class has:
- A static `namespace` property -- the `XNamespace` for that XML namespace URI.
- Static `XName` properties -- one for every element and attribute name in that namespace.

## Primary Namespace Classes

### W (WordprocessingML)

The `W` class covers WordprocessingML elements for Word documents.

```typescript
import { W } from "openxmlsdkts";

const body = xDoc.root!.element(W.body);
const paragraphs = body!.elements(W.p);

for (const para of paragraphs) {
  const runs = para.elements(W.r);
  for (const run of runs) {
    const text = run.element(W.t)?.value;
    console.log(text);
  }
}
```

### S (SpreadsheetML)

The `S` class covers SpreadsheetML elements for Excel workbooks.

```typescript
import { S } from "openxmlsdkts";

const sheetData = xDoc.root!.element(S.sheetData);
const rows = sheetData!.elements(S.row);

for (const row of rows) {
  const cells = row.elements(S.c);
  for (const cell of cells) {
    const value = cell.element(S.v)?.value;
    console.log(value);
  }
}
```

### P (PresentationML)

The `P` class covers PresentationML elements for PowerPoint presentations.

```typescript
import { P } from "openxmlsdkts";

const sldIdLst = xDoc.root!.element(P.sldIdLst);
const slideIds = sldIdLst!.elements(P.sldId);
console.log(`Presentation has ${slideIds.length} slides`);
```

### A (DrawingML)

The `A` class covers DrawingML elements shared across all three document types for drawings, shapes, charts, and graphics.

### R (Relationships)

The `R` class covers relationship elements used in `.rels` parts.

### NoNamespace

The `NoNamespace` class provides `XName` constants for unqualified attributes -- attributes that have no namespace prefix. These are extremely common in Open XML.

```typescript
import { NoNamespace, W } from "openxmlsdkts";

// Get the "val" attribute (no namespace) from a style element
const styleId = pStyle.attribute(NoNamespace.val)?.value;

// Get the "id" attribute
const id = element.attribute(NoNamespace.id)?.value;

// Get the "name" attribute
const name = element.attribute(NoNamespace.name)?.value;

// Get the "type" attribute
const type = element.attribute(NoNamespace.type)?.value;
```

## Full Reference Table

The following table lists all 48 namespace classes with their XML namespace URIs:

| Class | Namespace URI |
|-------|--------------|
| `A` | `http://schemas.openxmlformats.org/drawingml/2006/main` |
| `A14` | `http://schemas.microsoft.com/office/drawing/2010/main` |
| `C` | `http://schemas.openxmlformats.org/drawingml/2006/chart` |
| `CDR` | `http://schemas.openxmlformats.org/drawingml/2006/chartDrawing` |
| `COM` | `http://schemas.openxmlformats.org/drawingml/2006/compatibility` |
| `CP` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` |
| `CT` | `http://schemas.openxmlformats.org/package/2006/content-types` |
| `CUSTPRO` | `http://schemas.openxmlformats.org/officeDocument/2006/custom-properties` |
| `DC` | `http://purl.org/dc/elements/1.1/` |
| `DCTERMS` | `http://purl.org/dc/terms/` |
| `DGM` | `http://schemas.openxmlformats.org/drawingml/2006/diagram` |
| `DGM14` | `http://schemas.microsoft.com/office/drawing/2010/diagram` |
| `DIGSIG` | `http://schemas.microsoft.com/office/2006/digsig` |
| `DS` | `http://schemas.openxmlformats.org/officeDocument/2006/customXml` |
| `DSP` | `http://schemas.microsoft.com/office/drawing/2008/diagram` |
| `EP` | `http://schemas.openxmlformats.org/officeDocument/2006/extended-properties` |
| `FLATOPC` | `http://schemas.microsoft.com/office/2006/xmlPackage` |
| `LC` | `http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas` |
| `M` | `http://schemas.openxmlformats.org/officeDocument/2006/math` |
| `MC` | `http://schemas.openxmlformats.org/markup-compatibility/2006` |
| `MDSSI` | `http://schemas.openxmlformats.org/package/2006/digital-signature` |
| `MP` | `http://schemas.microsoft.com/office/mac/powerpoint/2008/main` |
| `MV` | `urn:schemas-microsoft-com:mac:vml` |
| `NoNamespace` | *(no namespace -- unqualified attributes)* |
| `O` | `urn:schemas-microsoft-com:office:office` |
| `P` | `http://schemas.openxmlformats.org/presentationml/2006/main` |
| `P14` | `http://schemas.microsoft.com/office/powerpoint/2010/main` |
| `P15` | `http://schemas.microsoft.com/office15/powerpoint` |
| `Pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` |
| `PKGREL` | `http://schemas.openxmlformats.org/package/2006/relationships` |
| `R` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| `S` | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` |
| `SL` | `http://schemas.openxmlformats.org/schemaLibrary/2006/main` |
| `SLE` | `http://schemas.microsoft.com/office/drawing/2010/slicer` |
| `VML` | `urn:schemas-microsoft-com:vml` |
| `VT` | `http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes` |
| `W` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` |
| `W10` | `urn:schemas-microsoft-com:office:word` |
| `W14` | `http://schemas.microsoft.com/office/word/2010/wordml` |
| `W3DIGSIG` | `http://www.w3.org/2000/09/xmldsig#` |
| `WP` | `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing` |
| `WP14` | `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing` |
| `WPS` | `http://schemas.microsoft.com/office/word/2010/wordprocessingShape` |
| `X` | `urn:schemas-microsoft-com:office:excel` |
| `XDR` | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` |
| `XDR14` | `http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing` |
| `XM` | `http://schemas.microsoft.com/office/excel/2006/main` |
| `XSI` | `http://www.w3.org/2001/XMLSchema-instance` |
