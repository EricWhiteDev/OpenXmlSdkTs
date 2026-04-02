/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

/**
 * Example: Load, Modify, and Save a Flat OPC File
 *
 * This example demonstrates:
 * - Opening a binary DOCX file and converting it to Flat OPC XML format
 * - Writing the Flat OPC XML to a file for inspection
 * - Reopening the document from the Flat OPC XML string
 * - Modifying the document using pre-atomized namespace names
 * - Saving the modified document as both Flat OPC XML and binary DOCX
 *
 * Flat OPC is an XML representation of an Office Open XML package where all
 * parts are embedded as <pkg:part> elements in a single XML document. This
 * format is useful for scenarios like Office Add-ins where binary ZIP files
 * cannot be used directly.
 */

import { WmlPackage, W, XElement } from "openxmlsdkts";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

(async () => {
  const exampleDir = path.resolve(__dirname, "example-output");
  fs.mkdirSync(exampleDir, { recursive: true });

  // Step 1: Open the source DOCX file as a binary Blob
  const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
  const buffer = fs.readFileSync(srcFile);
  const blob = new Blob([buffer]);
  const doc = await WmlPackage.open(blob);
  console.log("Opened TemplateDocument.docx successfully.");

  // Step 2: Convert to Flat OPC XML format
  // saveToFlatOpcAsync() serializes the entire package (all parts, relationships,
  // and content types) into a single XML string.
  const flatOpc = await doc.saveToFlatOpcAsync();
  const flatOpcOutputFile = path.join(exampleDir, "TemplateDocument.xml");
  fs.writeFileSync(flatOpcOutputFile, flatOpc, "utf-8");
  console.log(`Saved Flat OPC XML to: ${flatOpcOutputFile}`);
  console.log(`Flat OPC XML length: ${flatOpc.length} characters`);

  // Step 3: Reopen the document from the Flat OPC XML string
  // WmlPackage.open() auto-detects the format (binary Blob, base64 string, or
  // Flat OPC XML string) and handles each appropriately.
  const doc2 = await WmlPackage.open(flatOpc);
  console.log("Reopened document from Flat OPC XML.");

  // Step 4: Navigate and modify the document using pre-atomized names
  const mainPart = await doc2.mainDocumentPart();
  if (!mainPart) {
    throw new Error("Main document part not found.");
  }

  const xDoc = await mainPart.getXDocument();

  // Pre-atomized names provide type-safe, efficient XML navigation:
  //   W.body   => <w:body>     (the document body)
  //   W.p      => <w:p>        (paragraph)
  //   W.r      => <w:r>        (run - a contiguous piece of text with same formatting)
  //   W.t      => <w:t>        (text content)
  //   W.sectPr => <w:sectPr>   (section properties - page size, margins, etc.)
  const body = xDoc.root!.element(W.body);
  if (!body) {
    throw new Error("Body element not found.");
  }

  const paragraphs = Array.from(body.elements(W.p));
  console.log(`Document contains ${paragraphs.length} paragraphs.`);

  // Access the styles part to show part navigation
  const stylesPart = await mainPart.styleDefinitionsPart();
  if (stylesPart) {
    const stylesXDoc = await stylesPart.getXDocument();
    // W.style is the pre-atomized name for <w:style> elements
    const styles = Array.from(stylesXDoc.root!.elements(W.style));
    console.log(`Document defines ${styles.length} styles.`);
  }

  // Add a paragraph before the section properties
  const sectPr = body.element(W.sectPr);
  if (sectPr) {
    const newParagraph = new XElement(
      W.p,
      new XElement(
        W.r,
        new XElement(W.t, "This paragraph was added from a Flat OPC document."),
      ),
    );
    sectPr.addBeforeSelf(newParagraph);
    console.log("Added a new paragraph to the document.");
  }

  // Write the modified XML back to the part
  mainPart.putXDocument(xDoc);

  // Step 5: Save the modified document as Flat OPC XML
  const modifiedFlatOpc = await doc2.saveToFlatOpcAsync();
  const modifiedFlatOpcFile = path.join(exampleDir, "TemplateDocument-modified.xml");
  fs.writeFileSync(modifiedFlatOpcFile, modifiedFlatOpc, "utf-8");
  console.log(`Saved modified Flat OPC XML to: ${modifiedFlatOpcFile}`);

  // Step 6: Also save as binary DOCX
  const savedBlob = await doc2.saveToBlobAsync();
  const arrayBuffer = await savedBlob.arrayBuffer();
  const binaryOutputFile = path.join(exampleDir, "TemplateDocument-modified.docx");
  fs.writeFileSync(binaryOutputFile, Buffer.from(arrayBuffer));
  console.log(`Saved modified binary DOCX to: ${binaryOutputFile}`);

  console.log("Done!");
})().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
