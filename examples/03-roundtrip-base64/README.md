# Example: Round-Trip a DOCX Through Base64 Encoding

This example demonstrates round-tripping a Word document through base64 encoding, verifying that the document's content and structure are preserved at each step.

## What It Does

1. Opens `TemplateDocument.docx` as a binary DOCX file
2. Counts the paragraphs in the document body using **pre-atomized names** (`W.body`, `W.p`)
3. Saves the document as a **base64-encoded string** using `saveToBase64Async()`
4. Writes the base64 string to a text file for inspection
5. Reopens the document from the base64 string (auto-detected by `WmlPackage.open()`)
6. Verifies the paragraph count matches the original
7. Saves back to **binary DOCX** format
8. Reopens the saved DOCX and verifies document integrity one final time

## Why Base64?

Base64 encoding converts binary data to a text string using only ASCII characters. This is useful for:

- Transmitting documents over JSON APIs or other text-based protocols
- Embedding documents in data URIs
- Storing binary documents in text database fields
- Passing documents between systems that only support text data

## Pre-Atomized Namespaces

This example uses `W.body` and `W.p` to navigate the document XML. These pre-atomized `XName` objects correspond to the `<w:body>` and `<w:p>` WordprocessingML elements. Using these typed constants provides:

- **Type safety**: compile-time checking of element names
- **Performance**: identity-based equality (`===`) instead of string comparison
- **Readability**: clear, concise code without raw namespace URI strings

## Running

From the repository root:

```bash
node --conditions=import --import tsx Examples/03-roundtrip-base64/roundtrip-base64.ts
```

## Output

The example creates an `example-output/` directory containing:

- `TemplateDocument.b64.txt` -- the document as a base64-encoded string
- `TemplateDocument-roundtripped.docx` -- the document after the full round-trip
