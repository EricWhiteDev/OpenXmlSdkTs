---
title: "Getting Started"
group: "Guides"
category: "Getting Started"
---

# Getting Started

This guide walks you through installing OpenXmlSdkTs and performing your first document operations.

## Installation

```bash
npm install openxmlsdkts
```

The library has only two runtime dependencies (`jszip` and `ltxmlts`), which are installed automatically.

## Quick Start

Here is a complete example that opens a Word document, inspects it, adds a paragraph, and saves it:

```typescript
import { WmlPackage, W, XElement } from "openxmlsdkts";
import fs from "fs";

// Read the file into a Blob
const buffer = fs.readFileSync("report.docx");
const doc = await WmlPackage.open(new Blob([buffer]));

// Navigate to the main document part
const mainPart = await doc.mainDocumentPart();
const xDoc = await mainPart!.getXDocument();

// Query the XML
const body = xDoc.root!.element(W.body);
const paragraphs = body!.elements(W.p);
console.log(`Document has ${paragraphs.length} paragraphs`);

// Add a new paragraph
const newPara = new XElement(W.p,
  new XElement(W.r,
    new XElement(W.t, "Hello from OpenXmlSdkTs!")
  )
);
body!.add(newPara);

// Save changes back to the part and export
mainPart!.putXDocument(xDoc);
const blob = await doc.saveToBlobAsync();
```

## The Core Pattern

Every interaction with OpenXmlSdkTs follows the same five-step pattern:

1. **Import** -- Bring in the package class, namespace classes, and LINQ to XML types you need.
2. **Open** -- Call `WmlPackage.open()`, `SmlPackage.open()`, or `PmlPackage.open()` with a Blob, Flat OPC string, or Base64 string. The library auto-detects the format.
3. **Navigate** -- Use typed accessors to reach the part you want (e.g., `mainDocumentPart()`, `workbookPart()`).
4. **Query and Modify** -- Call `getXDocument()` to get the XML tree, then use LINQ to XML methods to read or change it. Call `putXDocument()` to write changes back to the part.
5. **Save** -- Export the modified package with `saveToBlobAsync()`, `saveToBase64Async()`, or `saveToFlatOpcAsync()`.

## Next Steps

- [Opening and Saving Documents](./03-opening-and-saving.md) -- Learn about the three I/O formats and how to convert between them.
- [Navigating Parts](./04-navigating-parts.md) -- Explore the full range of typed part accessors for Word, Excel, and PowerPoint.
- [Working with XML](./05-working-with-xml.md) -- Dive deeper into querying and constructing XML with LINQ to XML.
- [Namespace Classes](./06-namespaces.md) -- Understand pre-atomized names and the full set of namespace classes.
- [Office Add-ins](./08-office-add-ins.md) -- Use OpenXmlSdkTs inside Word, Excel, or PowerPoint add-ins.
