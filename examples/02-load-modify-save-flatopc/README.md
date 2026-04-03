# Example: Load, Modify, and Save a Flat OPC File

This example demonstrates working with the **Flat OPC** XML format, an alternative representation of Office Open XML packages where all parts are embedded in a single XML document.

## What It Does

1. Opens `TemplateDocument.docx` as a binary DOCX file
2. Converts the document to **Flat OPC XML** format using `saveToFlatOpcAsync()`
3. Writes the Flat OPC XML to a file so you can inspect the XML structure
4. Reopens the document from the Flat OPC XML string
5. Navigates the document structure using **pre-atomized namespace names**
6. Accesses the **styles part** and counts the defined styles
7. Adds a new paragraph to the document body
8. Saves the modified document as both **Flat OPC XML** and **binary DOCX**

## What is Flat OPC?

Flat OPC is an XML representation of an Office Open XML package. Instead of a ZIP file containing multiple XML files, all parts are embedded as `<pkg:part>` elements within a single XML document. This format is useful for:

- Office Add-ins (where binary ZIP files cannot be used directly)
- Storing documents in XML databases
- Transforming documents with XSLT
- Debugging (the entire document is human-readable in one file)

## Pre-Atomized Namespaces

This example uses pre-atomized names like `W.body`, `W.p`, `W.r`, `W.t`, and `W.style` to navigate the XML structure. These typed constants map to WordprocessingML elements and provide efficient identity-based equality checks. See the inline comments in the source code for details.

## Running

From the repository root:

```bash
node --conditions=import --import tsx Examples/02-load-modify-save-flatopc/load-modify-save-flatopc.ts
```

## Output

The example creates an `example-output/` directory containing:

- `TemplateDocument.xml` -- the original document in Flat OPC XML format
- `TemplateDocument-modified.xml` -- the modified document in Flat OPC XML format
- `TemplateDocument-modified.docx` -- the modified document as a binary DOCX file
