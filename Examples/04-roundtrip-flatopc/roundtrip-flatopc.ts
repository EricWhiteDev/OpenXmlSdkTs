/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

/**
 * Example: Round-Trip a DOCX Through Flat OPC XML
 *
 * This example demonstrates:
 * - Opening a binary DOCX file
 * - Saving the document as Flat OPC XML
 * - Reopening the document from the Flat OPC XML string
 * - Saving back to binary DOCX format
 * - Verifying that the round-trip preserves document integrity
 * - Using pre-atomized namespace names to inspect document content
 *
 * Flat OPC XML is a single-file XML representation of an Office Open XML
 * package, where all parts (document, styles, themes, etc.) are embedded
 * as <pkg:part> elements. This is the format used by Office Add-ins.
 */

import { WmlPackage, W } from "openxmlsdkts";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

(async () => {
  const exampleDir = path.resolve(__dirname, "example-output");
  fs.mkdirSync(exampleDir, { recursive: true });

  // Step 1: Open the original DOCX file
  const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
  const buffer = fs.readFileSync(srcFile);
  const blob = new Blob([buffer]);
  const doc = await WmlPackage.open(blob);
  console.log("Opened TemplateDocument.docx successfully.");

  // Step 2: Inspect the original document using pre-atomized names
  //
  // Pre-atomized names are pre-initialized XName objects that map to XML
  // element names in specific namespaces. For WordprocessingML:
  //   W.body    => {http://schemas.openxmlformats.org/wordprocessingml/2006/main}body
  //   W.p       => {http://schemas.openxmlformats.org/wordprocessingml/2006/main}p
  //   W.document => {http://schemas.openxmlformats.org/wordprocessingml/2006/main}document
  //
  // These objects are atomized: two references to W.p are the same object,
  // enabling O(1) identity comparison instead of string comparison.
  const mainPart = await doc.mainDocumentPart();
  if (!mainPart) {
    throw new Error("Main document part not found.");
  }
  const xDoc = await mainPart.getXDocument();
  const body = xDoc.root!.element(W.body);
  const originalParagraphCount = Array.from(body!.elements(W.p)).length;
  console.log(`Original document has ${originalParagraphCount} paragraphs.`);

  // Step 3: Save to Flat OPC XML
  // saveToFlatOpcAsync() produces a single XML string containing all parts
  // of the package, encoded as <pkg:part> elements.
  const flatOpc = await doc.saveToFlatOpcAsync();
  const flatOpcFile = path.join(exampleDir, "TemplateDocument.flatopc.xml");
  fs.writeFileSync(flatOpcFile, flatOpc, "utf-8");
  console.log(`Saved Flat OPC XML to: ${flatOpcFile}`);
  console.log(`Flat OPC XML length: ${flatOpc.length} characters`);

  // Step 4: Reload from Flat OPC XML
  // WmlPackage.open() auto-detects that this is a Flat OPC XML string
  // (it starts with "<?xml" or "<") and parses it accordingly.
  const doc2 = await WmlPackage.open(flatOpc);
  console.log("Reopened document from Flat OPC XML.");

  // Step 5: Verify the round-trip preserved the content
  const mainPart2 = await doc2.mainDocumentPart();
  if (!mainPart2) {
    throw new Error("Main document part not found after Flat OPC round-trip.");
  }
  const xDoc2 = await mainPart2.getXDocument();
  const body2 = xDoc2.root!.element(W.body);
  const flatOpcParagraphCount = Array.from(body2!.elements(W.p)).length;
  console.log(`After Flat OPC round-trip: ${flatOpcParagraphCount} paragraphs.`);

  if (originalParagraphCount !== flatOpcParagraphCount) {
    throw new Error(`Paragraph count mismatch: original=${originalParagraphCount}, after Flat OPC=${flatOpcParagraphCount}`);
  }
  console.log("Flat OPC round-trip: paragraph count matches.");

  // Step 6: Save back to binary DOCX
  const savedBlob = await doc2.saveToBlobAsync();
  const arrayBuffer = await savedBlob.arrayBuffer();
  const outputFile = path.join(exampleDir, "TemplateDocument-roundtripped.docx");
  fs.writeFileSync(outputFile, Buffer.from(arrayBuffer));
  console.log(`Saved round-tripped DOCX to: ${outputFile}`);

  // Step 7: Final verification - reopen the saved DOCX and check integrity
  const verifyBuffer = fs.readFileSync(outputFile);
  const verifyBlob = new Blob([verifyBuffer]);
  const verifyDoc = await WmlPackage.open(verifyBlob);
  const verifyPart = await verifyDoc.mainDocumentPart();
  if (!verifyPart) {
    throw new Error("Main document part not found in final verification.");
  }
  const verifyXDoc = await verifyPart.getXDocument();
  const verifyBody = verifyXDoc.root!.element(W.body);
  const finalParagraphCount = Array.from(verifyBody!.elements(W.p)).length;
  console.log(`Final verification: ${finalParagraphCount} paragraphs.`);

  if (originalParagraphCount !== finalParagraphCount) {
    throw new Error(`Final paragraph count mismatch: original=${originalParagraphCount}, final=${finalParagraphCount}`);
  }

  console.log("Full round-trip (DOCX -> Flat OPC -> DOCX) completed successfully!");
  console.log("Done!");
})().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
