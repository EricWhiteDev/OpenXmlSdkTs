/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

/**
 * Example: Round-Trip a DOCX Through Base64 Encoding
 *
 * This example demonstrates:
 * - Opening a binary DOCX file
 * - Saving the document as a base64-encoded string
 * - Reopening the document from the base64 string
 * - Saving back to binary DOCX format
 * - Verifying that the round-trip preserves document integrity
 * - Using pre-atomized namespace names to inspect document content
 *
 * Base64 encoding is useful for transmitting binary documents over text-based
 * protocols (e.g., JSON APIs, email, data URIs) or storing them in text fields.
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
  // W.body, W.p are pre-atomized XName objects for <w:body> and <w:p> elements.
  // They enable fast, type-safe XML queries without string parsing.
  const mainPart = await doc.mainDocumentPart();
  if (!mainPart) {
    throw new Error("Main document part not found.");
  }
  const xDoc = await mainPart.getXDocument();
  const body = xDoc.root!.element(W.body);
  const originalParagraphCount = Array.from(body!.elements(W.p)).length;
  console.log(`Original document has ${originalParagraphCount} paragraphs.`);

  // Step 3: Save to base64
  // saveToBase64Async() produces a base64-encoded string of the binary DOCX ZIP.
  const base64 = await doc.saveToBase64Async();
  const base64File = path.join(exampleDir, "TemplateDocument.b64.txt");
  fs.writeFileSync(base64File, base64, "utf-8");
  console.log(`Saved base64 to: ${base64File}`);
  console.log(`Base64 string length: ${base64.length} characters`);

  // Step 4: Reload from base64
  // WmlPackage.open() auto-detects that this is a base64 string and decodes it.
  const doc2 = await WmlPackage.open(base64);
  console.log("Reopened document from base64 string.");

  // Step 5: Verify the round-trip preserved the content
  const mainPart2 = await doc2.mainDocumentPart();
  if (!mainPart2) {
    throw new Error("Main document part not found after base64 round-trip.");
  }
  const xDoc2 = await mainPart2.getXDocument();
  const body2 = xDoc2.root!.element(W.body);
  const base64ParagraphCount = Array.from(body2!.elements(W.p)).length;
  console.log(`After base64 round-trip: ${base64ParagraphCount} paragraphs.`);

  if (originalParagraphCount !== base64ParagraphCount) {
    throw new Error(`Paragraph count mismatch: original=${originalParagraphCount}, after base64=${base64ParagraphCount}`);
  }
  console.log("Base64 round-trip: paragraph count matches.");

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

  console.log("Full round-trip (DOCX -> base64 -> DOCX) completed successfully!");
  console.log("Done!");
})().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
