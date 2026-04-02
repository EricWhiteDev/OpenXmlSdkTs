# Example: Round-Trip a DOCX Through Flat OPC XML

This example demonstrates round-tripping a Word document through the Flat OPC XML format, verifying that the document's content and structure are preserved at each step.

## What It Does

1. Opens `TemplateDocument.docx` as a binary DOCX file
2. Counts the paragraphs in the document body using **pre-atomized names** (`W.body`, `W.p`)
3. Saves the document as **Flat OPC XML** using `saveToFlatOpcAsync()`
4. Writes the Flat OPC XML to a file for inspection
5. Reopens the document from the Flat OPC XML string (auto-detected by `WmlPackage.open()`)
6. Verifies the paragraph count matches the original
7. Saves back to **binary DOCX** format
8. Reopens the saved DOCX and verifies document integrity one final time

## What is Flat OPC?

Flat OPC is a single-file XML representation of an Office Open XML package. All parts (document content, styles, themes, relationships, etc.) are embedded as `<pkg:part>` elements within one XML document. Key use cases:

- **Office Add-ins**: the Office JavaScript API uses Flat OPC for document manipulation
- **XML databases**: store complete documents as single XML entries
- **XSLT transformations**: transform documents using XML tooling
- **Debugging**: view the entire document structure in a single human-readable file

## Pre-Atomized Namespaces

This example uses `W.body`, `W.p`, and `W.document` to navigate the document XML. These are pre-atomized `XName` objects -- pre-initialized constants that map to the full namespace URI and local name of each WordprocessingML element:

```
W.body => {http://schemas.openxmlformats.org/wordprocessingml/2006/main}body
W.p    => {http://schemas.openxmlformats.org/wordprocessingml/2006/main}p
```

Because they are atomized (interned), two references to `W.p` are the exact same object in memory, enabling O(1) identity comparison.

## Running

From the repository root:

```bash
node --conditions=import --import tsx Examples/04-roundtrip-flatopc/roundtrip-flatopc.ts
```

## Output

The example creates an `example-output/` directory containing:

- `TemplateDocument.flatopc.xml` -- the document in Flat OPC XML format
- `TemplateDocument-roundtripped.docx` -- the document after the full round-trip
