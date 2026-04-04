---
title: "Navigating Parts"
group: "Guides"
category: "Core Concepts"
---

# Navigating Parts

An Open XML document is a ZIP package containing multiple XML parts connected by relationships. OpenXmlSdkTs provides typed accessors to navigate this structure.

## Package and Part Hierarchy

Every Open XML document follows the same structural pattern:

```
Package (the .docx/.xlsx/.pptx file)
├── Part (XML content)
│   ├── Relationship → Part
│   ├── Relationship → Part
│   └── ...
├── Part
│   └── Relationship → Part
└── ...
```

A **package** contains **parts**, and parts are connected to other parts through **relationships**. Each part has a content type (identifying what kind of XML it contains) and a URI (its path within the ZIP).

## Word Document Navigation (WmlPackage)

WmlPackage provides typed accessors for common Word document parts:

```typescript
const doc = await WmlPackage.open(blob);

// Main content
const mainPart = await doc.mainDocumentPart();

// Headers and footers
const headers = await doc.headerParts();
const footers = await doc.footerParts();

// Content parts (main document + headers + footers)
const allContent = await doc.contentParts();

// Styles and formatting
const styles = await doc.styleDefinitionsPart();

// Other common parts
const numbering = await doc.numberingDefinitionsPart();
const footnotes = await doc.footnotesPart();
const endnotes = await doc.endnotesPart();
const comments = await doc.commentsPart();
const settings = await doc.documentSettingsPart();
```

## Excel Workbook Navigation (SmlPackage)

SmlPackage provides typed accessors for spreadsheet parts:

```typescript
const workbook = await SmlPackage.open(blob);

// Workbook part (the root of spreadsheet content)
const wbPart = await workbook.workbookPart();

// Individual worksheets
const sheets = await workbook.worksheetParts();

// Shared strings (text values referenced by cells)
const sst = await workbook.sharedStringTablePart();

// Styles
const styles = await workbook.workbookStylesPart();
```

## PowerPoint Presentation Navigation (PmlPackage)

PmlPackage provides typed accessors for presentation parts:

```typescript
const pres = await PmlPackage.open(blob);

// Presentation part (the root of presentation content)
const presPart = await pres.presentationPart();

// Slides
const slides = await pres.slideParts();

// Slide masters and layouts
const masters = await pres.slideMasterParts();
const layouts = await pres.slideLayoutParts();
```

## Generic Relationship Navigation

When the typed accessors are not sufficient, you can navigate parts using content types and relationship types directly:

```typescript
// Get parts by relationship type
const parts = await mainPart!.getPartsByRelationshipType(
  RelationshipType.Image
);

// Get a single part by relationship type
const part = await mainPart!.getPartByRelationshipType(
  RelationshipType.StyleDefinitions
);

// Get parts by content type
const xmlParts = await mainPart!.getPartsByContentType(
  ContentType.WmlStyles
);
```

## Relationship Queries

You can query the relationships of any package or part:

```typescript
// Get all relationships
const rels = await mainPart!.getRelationships();

// Get relationships filtered by type
const imageRels = await mainPart!.getRelationshipsByRelationshipType(
  RelationshipType.Image
);

// Get a specific relationship by ID
const rel = await mainPart!.getRelationshipById("rId1");
```

Each `OpenXmlRelationship` provides:
- `id` -- The relationship ID (e.g., "rId1")
- `relationshipType` -- The relationship type URI
- `targetUri` -- The target part URI

## Part Management

You can add new parts to and remove existing parts from a package:

```typescript
// Add a new part
await mainPart!.addPart(newPart, relationshipType);

// Delete a part
await mainPart!.deletePart(existingPart);
```

## Property Parts

All package types share accessors for common metadata parts:

```typescript
// Core file properties (title, author, dates, etc.)
const corePropsPart = await doc.coreFilePropertiesPart();

// Extended file properties (application, company, etc.)
const extPropsPart = await doc.extendedFilePropertiesPart();
```

## Working with Part Content

Once you have navigated to a part, use `getXDocument()` and `putXDocument()` to read and write its XML content:

```typescript
const mainPart = await doc.mainDocumentPart();
const xDoc = await mainPart!.getXDocument();

// Query or modify xDoc...

mainPart!.putXDocument(xDoc);
```

See [Working with XML](./05-working-with-xml.md) for details on querying and modifying XML content.
