---
title: "Opening and Saving Documents"
group: "Guides"
category: "Core Concepts"
---

# Opening and Saving Documents

OpenXmlSdkTs supports three I/O formats for opening and saving documents. The `open()` method auto-detects which format you provide, so you can use whichever is most convenient for your scenario.

## Three I/O Formats

### Binary Blob

The standard approach for file system I/O. Use `Blob` (or `Buffer` in Node.js) to read and write `.docx`, `.xlsx`, and `.pptx` files directly.

```typescript
import { WmlPackage } from "openxmlsdkts";
import fs from "fs";

// Open from Blob
const buffer = fs.readFileSync("document.docx");
const doc = await WmlPackage.open(new Blob([buffer]));

// Save to Blob
const blob = await doc.saveToBlobAsync();
```

The type alias `DocxBinary` (and corresponding `XlsxBinary`, `PptxBinary`) represents binary document data.

### Flat OPC XML

Flat OPC is an XML representation of the entire Office document package in a single XML string. This format is especially useful for:

- **Office Add-ins** -- The Office JavaScript API uses Flat OPC for document exchange.
- **XML databases** -- Store documents as XML natively.
- **XSLT processing** -- Transform documents with standard XML tools.
- **Debugging** -- Inspect the full document structure as readable XML.

```typescript
import { WmlPackage } from "openxmlsdkts";

// Open from Flat OPC string
const flatOpc: string = getFlatOpcFromSomewhere();
const doc = await WmlPackage.open(flatOpc);

// Save to Flat OPC
const flatOpcOut = await doc.saveToFlatOpcAsync();
```

The type alias `FlatOpcString` represents Flat OPC XML data.

### Base64 String

A Base64-encoded representation of the binary package. Useful for:

- **JSON APIs** -- Embed document data in JSON payloads.
- **Data URIs** -- Use in browser contexts where data URIs are needed.
- **Text-based storage** -- Store in systems that only support text.

```typescript
import { WmlPackage } from "openxmlsdkts";

// Open from Base64 string
const base64: string = getBase64FromApi();
const doc = await WmlPackage.open(base64);

// Save to Base64
const base64Out = await doc.saveToBase64Async();
```

The type alias `Base64String` represents Base64-encoded document data.

## Auto-Detection

The `open()` method on each package class automatically detects the input format:

```typescript
// All three work with the same open() method
const doc1 = await WmlPackage.open(blob);       // Blob detected
const doc2 = await WmlPackage.open(flatOpc);     // Flat OPC XML detected
const doc3 = await WmlPackage.open(base64Str);   // Base64 detected
```

## Format-Specific Package Classes

Use the package class that matches your document type:

| Class | Document Type | Extension |
|-------|--------------|-----------|
| `WmlPackage` | Word documents | `.docx` |
| `SmlPackage` | Excel workbooks | `.xlsx` |
| `PmlPackage` | PowerPoint presentations | `.pptx` |

## Save Methods

Each package provides three save methods corresponding to the three I/O formats:

| Method | Returns | Description |
|--------|---------|-------------|
| `saveToBlobAsync()` | `Promise<Blob>` | Binary blob for file I/O |
| `saveToBase64Async()` | `Promise<string>` | Base64-encoded string |
| `saveToFlatOpcAsync()` | `Promise<string>` | Flat OPC XML string |

## Round-Trip Example

A typical workflow opens a document in one format, modifies it, and saves it in any format:

```typescript
import { WmlPackage, W, XElement } from "openxmlsdkts";

// Open from Base64 (e.g., received from an API)
const doc = await WmlPackage.open(base64Input);

// Modify the document
const mainPart = await doc.mainDocumentPart();
const xDoc = await mainPart!.getXDocument();
const body = xDoc.root!.element(W.body);
body!.add(
  new XElement(W.p,
    new XElement(W.r,
      new XElement(W.t, "Added by OpenXmlSdkTs")
    )
  )
);
mainPart!.putXDocument(xDoc);

// Save as Flat OPC (e.g., to send to an Office Add-in)
const flatOpc = await doc.saveToFlatOpcAsync();
```

You can open in any format and save in any other format -- the library handles all conversions internally.
