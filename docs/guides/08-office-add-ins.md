---
title: "Office Add-ins"
group: "Guides"
category: "Scenarios"
---

# Office Add-ins

OpenXmlSdkTs works in Office Add-ins (also called Office Web Add-ins), allowing you to read and modify the current document directly from within Word, Excel, or PowerPoint. The key to this integration is the **Flat OPC** format -- the XML representation that the Office JavaScript API uses for document exchange.

## Why Flat OPC?

The Office JavaScript API provides access to the current document's content as a Flat OPC XML string. Flat OPC represents the entire Open XML package (all parts, relationships, and content types) as a single XML document. This is the bridge between the Office host application and OpenXmlSdkTs.

## Getting the Document from Office.js

The Office JavaScript API provides methods to retrieve the document content. The exact approach depends on the host application and API set you are targeting.

### Word Add-ins

In Word, you can use the `getOoxml()` method to retrieve content as Office Open XML:

```typescript
await Word.run(async (context) => {
  const body = context.document.body;
  const ooxml = body.getOoxml();
  await context.sync();

  // ooxml.value contains the Flat OPC XML
  const flatOpc = ooxml.value;
});
```

### Using getFileAsync

For all Office hosts, `Office.context.document.getFileAsync` can retrieve the full document:

```typescript
Office.context.document.getFileAsync(
  Office.FileType.Compressed,
  { sliceSize: 65536 },
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const file = result.value;
      // Read slices and assemble the document...
    }
  }
);
```

## Converting Between Formats

OpenXmlSdkTs handles format conversion seamlessly. The `open()` method auto-detects whether you pass a Blob, Flat OPC string, or Base64 string:

```typescript
import { WmlPackage } from "openxmlsdkts";

// Open from Flat OPC (received from Office.js)
const doc = await WmlPackage.open(flatOpcString);

// Save back to Flat OPC (to send back to Office.js)
const flatOpcOut = await doc.saveToFlatOpcAsync();

// Or convert to other formats as needed
const blob = await doc.saveToBlobAsync();
const base64 = await doc.saveToBase64Async();
```

## Example Workflow

Here is a complete example of an Office Add-in workflow that opens the current Word document, adds a paragraph, and writes it back:

```typescript
import { WmlPackage, W, XElement } from "openxmlsdkts";

async function addParagraphToDocument(): Promise<void> {
  await Word.run(async (context) => {
    // Step 1: Get the document as Flat OPC from Office.js
    const body = context.document.body;
    const ooxml = body.getOoxml();
    await context.sync();
    const flatOpc = ooxml.value;

    // Step 2: Open with OpenXmlSdkTs
    const doc = await WmlPackage.open(flatOpc);
    const mainPart = await doc.mainDocumentPart();
    const xDoc = await mainPart!.getXDocument();

    // Step 3: Modify the document
    const docBody = xDoc.root!.element(W.body);
    const newParagraph = new XElement(W.p,
      new XElement(W.r,
        new XElement(W.rPr,
          new XElement(W.b)
        ),
        new XElement(W.t, "Added by Office Add-in using OpenXmlSdkTs")
      )
    );
    docBody!.add(newParagraph);
    mainPart!.putXDocument(xDoc);

    // Step 4: Save back to Flat OPC
    const modifiedFlatOpc = await doc.saveToFlatOpcAsync();

    // Step 5: Write back to the document via Office.js
    body.insertOoxml(modifiedFlatOpc, Word.InsertLocation.replace);
    await context.sync();
  });
}
```

## Browser Environment Considerations

When running in an Office Add-in (which executes in a browser or webview), keep the following in mind:

- **No file system access** -- You cannot use `fs.readFileSync` or other Node.js APIs. Use the Office JavaScript API to get document content, and use Blob or Base64 formats for any external data.
- **Bundling** -- OpenXmlSdkTs and its dependencies (`jszip`, `ltxmlts`) need to be bundled for browser use. Standard bundlers like webpack, esbuild, or Vite work well.
- **Async operations** -- All package open and save operations are asynchronous. Make sure to `await` them properly within the `Word.run()`, `Excel.run()`, or `PowerPoint.run()` callback.
- **Memory** -- Large documents consume memory in the browser. Be mindful of document size, especially on mobile Office clients.
- **CORS** -- If your add-in fetches documents from external services, ensure proper CORS headers are configured on the server.
