# Example: Load, Modify, and Save a Binary DOCX File

This example demonstrates the most common workflow with OpenXmlSdkTs: opening a binary DOCX file, navigating its XML structure, modifying content, and saving the result.

## What It Does

1. Opens `WithComments.docx` as a binary DOCX file
2. Navigates to the main document part and accesses its XML
3. Counts the paragraphs in the document body
4. Accesses the **comments part** and reads each comment's author and text
5. Adds a new **bold paragraph** to the end of the document body
6. Saves the modified document back to DOCX format
7. Reopens the saved file to verify it was written correctly

## Pre-Atomized Namespaces

This example highlights the use of **pre-atomized namespace names** -- a key feature of OpenXmlSdkTs. Instead of working with raw XML namespace URIs and element names as strings, you use typed constants:

- `W.body` corresponds to `<w:body>` (the document body element)
- `W.p` corresponds to `<w:p>` (a paragraph element)
- `W.r` corresponds to `<w:r>` (a run element)
- `W.t` corresponds to `<w:t>` (a text element)
- `W.rPr` corresponds to `<w:rPr>` (run properties)
- `W.b` corresponds to `<w:b>` (bold formatting)
- `W.comment` corresponds to `<w:comment>` (a comment element)

These are pre-initialized `XName` objects. Because they are atomized (interned), equality checks use fast identity comparison (`===`) rather than string comparison.

## Running

From the repository root:

```bash
node --conditions=import --import tsx Examples/01-load-modify-save-binary/load-modify-save-binary.ts
```

## Output

The example creates an `example-output/` directory containing:

- `WithComments-modified.docx` -- the modified document with the added paragraph
