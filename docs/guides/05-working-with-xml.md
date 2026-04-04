---
title: "Working with XML"
group: "Guides"
category: "Core Concepts"
---

# Working with XML

OpenXmlSdkTs uses LINQ to XML (provided by the [ltxmlts](https://www.npmjs.com/package/ltxmlts) package) for all XML operations. This gives you a powerful, functional API for querying, constructing, and modifying XML trees.

## The getXDocument / putXDocument Pattern

Every part's XML content is accessed through a simple read/write cycle:

```typescript
// Read the XML tree from the part
const xDoc = await part.getXDocument();

// Query or modify the tree...

// Write the modified tree back to the part
part.putXDocument(xDoc);
```

`getXDocument()` returns an `XDocument` -- the root of a LINQ to XML tree. After making changes, call `putXDocument()` to persist them. Changes are not saved to the package until you call one of the save methods on the package itself.

## LINQ to XML Core Types

| Type | Description |
|------|-------------|
| `XDocument` | Represents an XML document; has a `root` property pointing to the root element |
| `XElement` | An XML element with a name, attributes, and child nodes |
| `XAttribute` | A name-value pair on an element |
| `XName` | A qualified XML name (namespace + local name) |
| `XNamespace` | An XML namespace URI |

## Querying XML

### Finding Elements

```typescript
import { W } from "openxmlsdkts";

const xDoc = await mainPart!.getXDocument();
const root = xDoc.root!;

// Get the first child element with a specific name
const body = root.element(W.body);

// Get all child elements with a specific name
const paragraphs = body!.elements(W.p);

// Get all child elements (regardless of name)
const allChildren = body!.elements();

// Get all descendant elements with a specific name (searches the full subtree)
const allRuns = root.descendants(W.r);
```

### Reading Attributes

```typescript
import { W, NoNamespace } from "openxmlsdkts";

// Get an attribute by name
const valAttr = element.attribute(NoNamespace.val);

// Get the attribute value as a string
const value = valAttr?.value;

// Get the value of the w:val attribute on a style element
const styleId = styleElement.attribute(NoNamespace.val)?.value;
```

### Chaining Queries

Queries can be chained to drill into the document structure:

```typescript
// Get all text content from a document
const texts = xDoc.root!
  .element(W.body)!
  .elements(W.p)
  .flatMap(p => p.elements(W.r))
  .flatMap(r => r.elements(W.t))
  .map(t => t.value);
```

## Constructing XML

Create new XML elements using the `XElement` constructor. The constructor accepts the element name followed by any combination of attributes, child elements, and text content:

```typescript
import { XElement, XAttribute } from "ltxmlts";
import { W, NoNamespace } from "openxmlsdkts";

// Create a simple paragraph with text
const paragraph = new XElement(W.p,
  new XElement(W.r,
    new XElement(W.rPr,
      new XElement(W.b)  // bold
    ),
    new XElement(W.t, "Bold text here")
  )
);

// Create an element with attributes
const bookmark = new XElement(W.bookmarkStart,
  new XAttribute(NoNamespace.id, "0"),
  new XAttribute(NoNamespace.name, "MyBookmark")
);
```

## Modifying XML

### Adding Content

```typescript
// Add a child element at the end
body!.add(newParagraph);

// Add an element before a sibling
existingParagraph.addBeforeSelf(newParagraph);

// Add an element after a sibling
existingParagraph.addAfterSelf(newParagraph);
```

### Removing Content

```typescript
// Remove an element from its parent
paragraphToDelete.remove();
```

### Replacing Content

```typescript
// Replace an element with another
oldElement.replaceWith(newElement);
```

### Setting Attribute Values

```typescript
// Set or add an attribute
element.setAttributeValue(NoNamespace.val, "newValue");
```

## Example: Reading Paragraphs from a Word Document

```typescript
import { WmlPackage, W, NoNamespace } from "openxmlsdkts";

const doc = await WmlPackage.open(blob);
const mainPart = await doc.mainDocumentPart();
const xDoc = await mainPart!.getXDocument();

const body = xDoc.root!.element(W.body)!;

for (const para of body.elements(W.p)) {
  // Get paragraph style
  const pStyle = para.element(W.pPr)?.element(W.pStyle);
  const styleName = pStyle?.attribute(NoNamespace.val)?.value ?? "(default)";

  // Get concatenated text
  const text = para.elements(W.r)
    .flatMap(r => r.elements(W.t))
    .map(t => t.value)
    .join("");

  console.log(`[${styleName}] ${text}`);
}
```

## Example: Modifying Cell Values in an Excel Worksheet

```typescript
import { SmlPackage, S, NoNamespace } from "openxmlsdkts";

const workbook = await SmlPackage.open(blob);
const sheets = await workbook.worksheetParts();
const sheet1 = sheets[0];
const xDoc = await sheet1.getXDocument();

const sheetData = xDoc.root!.element(S.sheetData)!;

for (const row of sheetData.elements(S.row)) {
  for (const cell of row.elements(S.c)) {
    const ref = cell.attribute(NoNamespace.r)?.value;
    const value = cell.element(S.v)?.value;
    console.log(`Cell ${ref}: ${value}`);
  }
}

// Set a cell value
const targetCell = sheetData
  .elements(S.row)
  .flatMap(r => r.elements(S.c))
  .find(c => c.attribute(NoNamespace.r)?.value === "A1");

if (targetCell) {
  const vElement = targetCell.element(S.v);
  if (vElement) {
    vElement.value = "42";
  }
}

sheet1.putXDocument(xDoc);
```

## Further Reading

The LINQ to XML API is provided by the `ltxmlts` package. See its [npm page](https://www.npmjs.com/package/ltxmlts) for the full API reference, including advanced features like XML serialization, namespace declarations, and more.
