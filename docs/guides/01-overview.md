---
title: "Library Overview"
group: "Guides"
category: "Getting Started"
---

# Library Overview

OpenXmlSdkTs is a TypeScript library for reading, writing, and manipulating Office Open XML documents -- the file formats behind `.docx`, `.xlsx`, and `.pptx` files. It brings the programming model of the well-known .NET Open XML SDK to the TypeScript/JavaScript ecosystem.

## Key Features

- **Three I/O formats** -- Open and save documents as binary Blobs, Flat OPC XML strings, or Base64-encoded strings. The library auto-detects the format on open.
- **Pre-atomized namespace names** -- All XML element and attribute names are pre-atomized as static `XName` properties on dedicated namespace classes (W, S, P, A, R, NoNamespace, and more). This enables O(1) identity comparisons instead of string matching.
- **LINQ to XML** -- Built on [ltxmlts](https://www.npmjs.com/package/ltxmlts), a TypeScript port of .NET's LINQ to XML. Query and construct XML trees with `XDocument`, `XElement`, `XAttribute`, `XName`, and `XNamespace`.
- **Typed part navigation** -- Navigate document parts through strongly-typed accessors like `mainDocumentPart()`, `workbookPart()`, and `presentationPart()`.
- **Lightweight** -- Only two runtime dependencies: `jszip` and `ltxmlts`.

## Class Hierarchy

```
OpenXmlPackage (base)
├── WmlPackage (.docx)
├── SmlPackage (.xlsx)
└── PmlPackage (.pptx)

OpenXmlPart (base)
├── WmlPart (Word parts)
├── SmlPart (Excel parts)
└── PmlPart (PowerPoint parts)

OpenXmlRelationship
```

Packages contain parts, and parts contain XML content. You open a package, navigate to the part you need, retrieve its XML document, query or modify it, and save the package back out.

## Namespace Classes

The library provides over 48 namespace classes that give you pre-atomized `XName` constants for every element and attribute in the Open XML specification:

| Class | Description |
|-------|-------------|
| `W` | WordprocessingML (word processing elements) |
| `S` | SpreadsheetML (spreadsheet elements) |
| `P` | PresentationML (presentation elements) |
| `A` | DrawingML (drawing/graphics elements) |
| `R` | Relationships |
| `NoNamespace` | Unqualified attributes (val, id, type, name, etc.) |
| `M`, `MC`, `WP`, `WPS`, `WPC`, `W14`, `W15`, ... | Additional specialized namespaces |

Each class exposes a static `namespace` property (the `XNamespace`) and static `XName` properties for every element and attribute name in that namespace.

## ContentType and RelationshipType

The library includes two lookup objects that enumerate every content type and relationship type defined in the Open XML specification:

- **ContentType** -- Maps descriptive keys to MIME content type strings (e.g., `ContentType.WmlDocument` resolves to `application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml`).
- **RelationshipType** -- Maps descriptive keys to relationship type URIs (e.g., `RelationshipType.OfficeDocument` resolves to the appropriate URI).

These are used when querying for parts by type or when adding new parts and relationships to a package.

## Use Cases

- **Office Add-ins** -- Manipulate the current document inside Word, Excel, or PowerPoint add-ins using Flat OPC format.
- **Server-side document generation** -- Create or modify Office documents in Node.js services and APIs.
- **Document transformation** -- Convert, merge, split, or restructure documents programmatically.
- **Inspection and analysis** -- Extract text, metadata, styles, or structural information from documents.

## Inspiration

OpenXmlSdkTs is inspired by the [Open XML SDK for .NET](https://github.com/dotnet/Open-XML-SDK). Developers familiar with the .NET SDK will recognize the same package/part/relationship model and the same LINQ to XML approach to working with document content.
