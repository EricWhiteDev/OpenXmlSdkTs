/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { WmlPackage, W, ContentType, RelationshipType, XDocument } from "OpenXmlSdkTs";
import { blankDocumentBase64, blankDocumentFlatOpc } from "./TestResources";
import * as fs from "fs";
import * as path from "path";

describe("WmlPackage", () => {
  it("does not throw when opening a docx blob", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    await expect(WmlPackage.open(blob)).resolves.toBeDefined();
  });

  it("mainDocumentPart returns the main document part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const part = await doc.mainDocumentPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/document.xml");
  });

  it("contentParts returns main document, headers, and footers", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithPageHeaderAndPageFooter.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const parts = await doc.contentParts();
    const uris = parts.map((p) => p.getUri());
    expect(uris).toContain("/word/document.xml");
    expect(uris).toContain("/word/header1.xml");
    expect(uris).toContain("/word/footer1.xml");
    expect(parts[0].getUri()).toBe("/word/document.xml");
  });

  it("gets relationships for the document part from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const rels = await docPart.getRelationships();

    const commentsRel = rels.find((r) => r.getType() === RelationshipType.wordprocessingComments);
    expect(commentsRel).toBeDefined();
    expect(commentsRel!.getTarget()).toBe("comments.xml");
  });

  it("gets part-level relationships by type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const rels = await docPart.getRelationshipsByRelationshipType(RelationshipType.wordprocessingComments);

    expect(rels).toHaveLength(1);
    expect(rels[0].getTarget()).toBe("comments.xml");
  });

  it("gets part by relationship type from part level in WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.getPartByRelationshipType(RelationshipType.wordprocessingComments);

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/comments.xml");
  });

  it("returns undefined from getPartByRelationshipType when no match exists", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.getPartByRelationshipType(RelationshipType.wordprocessingComments);

    expect(part).toBeUndefined();
  });

  it("gets parts by relationship type from part level in WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const parts = await docPart.getPartsByRelationshipType(RelationshipType.wordprocessingComments);

    expect(parts).toHaveLength(1);
    expect(parts[0].getUri()).toBe("/word/comments.xml");
  });

  it("adds a part-level relationship and round-trips correctly", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const rel = await docPart.addRelationship("rId99", RelationshipType.wordprocessingComments, "comments.xml");
    expect(rel.getId()).toBe("rId99");
    expect(rel.getType()).toBe(RelationshipType.wordprocessingComments);
    expect(rel.getTarget()).toBe("comments.xml");
    expect(rel.getTargetMode()).toBeNull();

    const saved = await doc.saveToBase64Async();
    const doc2 = await WmlPackage.open(saved);
    const docPart2 = (await doc2.mainDocumentPart())!;
    const roundTrippedRel = await docPart2.getRelationshipById("rId99");
    expect(roundTrippedRel).toBeDefined();
    expect(roundTrippedRel!.getType()).toBe(RelationshipType.wordprocessingComments);
    expect(roundTrippedRel!.getTarget()).toBe("comments.xml");
  });

  it("gets a part-level relationship by id from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const rel = await docPart.getRelationshipById("rId4");

    expect(rel).toBeDefined();
    expect(rel!.getType()).toBe(RelationshipType.wordprocessingComments);
    expect(rel!.getTarget()).toBe("comments.xml");
  });

  it("gets a part-level part by id from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.getPartById("rId4");

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/comments.xml");
  });

  it("gets part-level relationships by content type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const rels = await docPart.getRelationshipsByContentType(ContentType.wordprocessingComments);

    expect(rels).toHaveLength(1);
    expect(rels[0].getTarget()).toBe("comments.xml");
  });

  it("gets part-level parts by content type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const parts = await docPart.getPartsByContentType(ContentType.wordprocessingComments);

    expect(parts).toHaveLength(1);
    expect(parts[0].getUri()).toBe("/word/comments.xml");
  });

  it("deletes a part-level relationship and verifies it is gone after round-trip", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const result = await docPart.deleteRelationship("rId4");
    expect(result).toBe(true);
    expect(await docPart.getRelationshipById("rId4")).toBeUndefined();

    const saved = await doc.saveToBase64Async();
    const doc2 = await WmlPackage.open(saved);
    const docPart2 = (await doc2.mainDocumentPart())!;
    expect(await docPart2.getRelationshipById("rId4")).toBeUndefined();
  });

  it("throws when deleting a part-level relationship that does not exist", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    await expect(docPart.deleteRelationship("rIdDoesNotExist")).rejects.toThrow("Relationship not found: rIdDoesNotExist");
  });

  it("getXDocument materializes a lazy blob-opened part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const xDoc = await docPart.getXDocument();
    expect(xDoc).toBeDefined();
    expect(xDoc.root).not.toBeNull();
    expect(xDoc.root!.element(W.body)).not.toBeNull();
  });

  it("getXDocument returns the XDocument directly for a FlatOPC-opened part", async () => {
    const doc = await WmlPackage.open(blankDocumentFlatOpc);
    const docPart = (await doc.mainDocumentPart())!;
    const xDoc = await docPart.getXDocument();
    expect(xDoc).toBeDefined();
    expect(xDoc.root!.element(W.body)).not.toBeNull();
  });

  it("getXDocument throws for a non-xml part", async () => {
    const doc = await WmlPackage.open(blankDocumentBase64);
    const imgPart = doc.addPart("/word/media/img.png", "image/png", "binary", "fakedata");
    await expect(imgPart.getXDocument()).rejects.toThrow("Cannot get XDocument for non-xml part: /word/media/img.png");
  });

  it("putXDocument throws when xDoc is null", async () => {
    const doc = await WmlPackage.open(blankDocumentBase64);
    const docPart = (await doc.mainDocumentPart())!;
    expect(() => docPart.putXDocument(null as unknown as XDocument)).toThrow("putXDocument: xDoc must not be null or undefined");
  });

  it("getParts returns related parts of document part from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const parts = await docPart.getParts();

    expect(parts.length).toBeGreaterThan(0);
    expect(parts.some((p) => p.getUri() === "/word/comments.xml")).toBe(true);
  });

  it("wordprocessingCommentsPart returns comments part from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.wordprocessingCommentsPart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/comments.xml");
  });

  it("styleDefinitionsPart returns styles part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.styleDefinitionsPart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/styles.xml");
  });

  it("documentSettingsPart returns settings part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.documentSettingsPart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/settings.xml");
  });

  it("fontTablePart returns fontTable part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.fontTablePart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/fontTable.xml");
  });

  it("webSettingsPart returns webSettings part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.webSettingsPart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/webSettings.xml");
  });

  it("themePart returns theme part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.themePart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/theme/theme1.xml");
  });

  it("calculationChainPart returns undefined on a Word document", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.calculationChainPart();

    expect(part).toBeUndefined();
  });

  it("sharedStringTablePart returns undefined on a Word document", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.sharedStringTablePart();

    expect(part).toBeUndefined();
  });

  it("worksheetParts returns empty array on a Word document", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const parts = await docPart.worksheetParts();

    expect(parts).toHaveLength(0);
  });

  it("slideParts returns empty array on a Word document", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const parts = await docPart.slideParts();

    expect(parts).toHaveLength(0);
  });

  it("imageParts returns empty array on a Word document with no images", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const parts = await docPart.imageParts();

    expect(parts).toHaveLength(0);
  });

  it("getParts throws for a dangling internal relationship", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    await docPart.addRelationship("rId99", RelationshipType.wordprocessingComments, "missing.xml");

    await expect(docPart.getParts()).rejects.toThrow("Part not found for relationship target: /word/missing.xml");
  });
});
