/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

/**
 * Example: Load, Modify, and Save a Binary DOCX File
 *
 * This example demonstrates:
 * - Opening a binary DOCX file using WmlPackage.open()
 * - Navigating to the main document part and accessing its XML
 * - Using pre-atomized namespace names (W.body, W.p, W.r, W.t, etc.)
 * - Accessing the comments part and reading comment data
 * - Adding a new bold paragraph to the document
 * - Saving the modified document back to a DOCX file
 */

import { WmlPackage, W, NoNamespace, XElement } from "openxmlsdkts";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

(async () => {
  const exampleDir = path.resolve(__dirname, "example-output");
  fs.mkdirSync(exampleDir, { recursive: true });

  // Copy the source document so we don't modify the original test file
  const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
  const outputFile = path.join(exampleDir, "WithComments-modified.docx");
  fs.copyFileSync(srcFile, outputFile);

  // Open the DOCX file as a binary Blob
  const buffer = fs.readFileSync(outputFile);
  const blob = new Blob([buffer]);
  const doc = await WmlPackage.open(blob);
  console.log("Opened WithComments.docx successfully.");

  // Navigate to the main document part
  const mainPart = await doc.mainDocumentPart();
  if (!mainPart) {
    throw new Error("Main document part not found.");
  }

  // Get the XML document (XDocument) for the main part
  const xDoc = await mainPart.getXDocument();

  // Use pre-atomized names to navigate the XML structure.
  //
  // Pre-atomized names like W.body, W.p, W.r, W.t are pre-initialized XName objects
  // that correspond to WordprocessingML elements:
  //   W.body  => <w:body>     (document body)
  //   W.p     => <w:p>        (paragraph)
  //   W.r     => <w:r>        (run)
  //   W.t     => <w:t>        (text)
  //   W.rPr   => <w:rPr>      (run properties)
  //   W.b     => <w:b>        (bold)
  //   W.sectPr => <w:sectPr>  (section properties)
  //
  // Because these are atomized, equality checks use identity comparison (===),
  // which is faster than string comparison.

  const body = xDoc.root!.element(W.body);
  if (!body) {
    throw new Error("Body element not found.");
  }

  // Count paragraphs using the pre-atomized W.p name
  const paragraphs = Array.from(body.elements(W.p));
  console.log(`Document contains ${paragraphs.length} paragraphs.`);

  // Access the comments part using the typed navigation method
  const commentsPart = await mainPart.wordprocessingCommentsPart();
  if (commentsPart) {
    const commentsXDoc = await commentsPart.getXDocument();

    // W.comment is the pre-atomized name for <w:comment> elements
    // NoNamespace.author is the pre-atomized name for the "author" attribute (no namespace)
    const comments = Array.from(commentsXDoc.root!.elements(W.comment));
    console.log(`Document contains ${comments.length} comment(s):`);
    for (const comment of comments) {
      const author = comment.attribute(W.author)?.value ?? "(unknown)";
      // Get the text content of the comment's paragraphs
      const textParts: string[] = [];
      for (const t of comment.descendants(W.t)) {
        textParts.push(t.value);
      }
      console.log(`  - Author: ${author}, Text: "${textParts.join("")}"`);
    }
  } else {
    console.log("No comments part found in this document.");
  }

  // Add a new bold paragraph before the section properties element (W.sectPr)
  const sectPr = body.element(W.sectPr);
  if (sectPr) {
    // Build a new paragraph using pre-atomized element names:
    //   W.p  -> paragraph
    //   W.r  -> run
    //   W.rPr -> run properties
    //   W.b  -> bold formatting
    //   W.t  -> text content
    const newParagraph = new XElement(
      W.p,
      new XElement(
        W.r,
        new XElement(W.rPr, new XElement(W.b)),
        new XElement(W.t, "This paragraph was added by the OpenXmlSdkTs example."),
      ),
    );
    sectPr.addBeforeSelf(newParagraph);
    console.log("Added a new bold paragraph to the document.");
  }

  // Write the modified XML back to the part
  mainPart.putXDocument(xDoc);

  // Save the modified document back to a DOCX file
  const savedBlob = await doc.saveToBlobAsync();
  const arrayBuffer = await savedBlob.arrayBuffer();
  fs.writeFileSync(outputFile, Buffer.from(arrayBuffer));

  console.log(`Saved modified document to: ${outputFile}`);

  // Verify: reopen the saved file and confirm it's valid
  const verifyBuffer = fs.readFileSync(outputFile);
  const verifyBlob = new Blob([verifyBuffer]);
  const verifyDoc = await WmlPackage.open(verifyBlob);
  const verifyPart = await verifyDoc.mainDocumentPart();
  const verifyXDoc = await verifyPart!.getXDocument();
  const verifyBody = verifyXDoc.root!.element(W.body);
  const verifyParagraphs = Array.from(verifyBody!.elements(W.p));
  console.log(`Verification: reopened document has ${verifyParagraphs.length} paragraphs (was ${paragraphs.length}).`);
  console.log("Done!");
})().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
