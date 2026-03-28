/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { OpenXmlPackage, W, ContentType, RelationshipType, XDocument } from "OpenXmlSdkTs";
import { blankDocumentBase64, blankDocumentFlatOpc } from "./TestResources";
import * as fs from "fs";
import * as path from "path";

describe("OpenXmlPart", () => {
  it("gets relationships for the document part from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const rels = await docPart.getRelationships();

    const commentsRel = rels.find((r) => r.getType() === RelationshipType.wordprocessingComments);
    expect(commentsRel).toBeDefined();
    expect(commentsRel!.getTarget()).toBe("comments.xml");
  });

  it("gets part-level relationships by type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const rels = await docPart.getRelationshipsByRelationshipType(
      RelationshipType.wordprocessingComments,
    );

    expect(rels).toHaveLength(1);
    expect(rels[0].getTarget()).toBe("comments.xml");
  });

  it("gets part by relationship type from part level in WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const part = await docPart.getPartByRelationshipType(RelationshipType.wordprocessingComments);

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/comments.xml");
  });

  it("returns undefined from getPartByRelationshipType when no match exists", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const part = await docPart.getPartByRelationshipType(RelationshipType.wordprocessingComments);

    expect(part).toBeUndefined();
  });

  it("gets parts by relationship type from part level in WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const parts = await docPart.getPartsByRelationshipType(RelationshipType.wordprocessingComments);

    expect(parts).toHaveLength(1);
    expect(parts[0].getUri()).toBe("/word/comments.xml");
  });

  it("adds a part-level relationship and round-trips correctly", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const rel = await docPart.addRelationship(
      "rId99",
      RelationshipType.wordprocessingComments,
      "comments.xml",
    );
    expect(rel.getId()).toBe("rId99");
    expect(rel.getType()).toBe(RelationshipType.wordprocessingComments);
    expect(rel.getTarget()).toBe("comments.xml");
    expect(rel.getTargetMode()).toBeNull();

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);
    const docPart2 = pkg2.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const roundTrippedRel = await docPart2.getRelationshipById("rId99");
    expect(roundTrippedRel).toBeDefined();
    expect(roundTrippedRel!.getType()).toBe(RelationshipType.wordprocessingComments);
    expect(roundTrippedRel!.getTarget()).toBe("comments.xml");
  });

  it("gets a part-level relationship by id from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const rel = await docPart.getRelationshipById("rId4");

    expect(rel).toBeDefined();
    expect(rel!.getType()).toBe(RelationshipType.wordprocessingComments);
    expect(rel!.getTarget()).toBe("comments.xml");
  });

  it("gets a part-level part by id from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const part = await docPart.getPartById("rId4");

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/comments.xml");
  });

  it("gets part-level relationships by content type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const rels = await docPart.getRelationshipsByContentType(ContentType.wordprocessingComments);

    expect(rels).toHaveLength(1);
    expect(rels[0].getTarget()).toBe("comments.xml");
  });

  it("gets part-level parts by content type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const parts = await docPart.getPartsByContentType(ContentType.wordprocessingComments);

    expect(parts).toHaveLength(1);
    expect(parts[0].getUri()).toBe("/word/comments.xml");
  });

  it("deletes a part-level relationship and verifies it is gone after round-trip", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const result = await docPart.deleteRelationship("rId4");
    expect(result).toBe(true);
    expect(await docPart.getRelationshipById("rId4")).toBeUndefined();

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);
    const docPart2 = pkg2.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    expect(await docPart2.getRelationshipById("rId4")).toBeUndefined();
  });

  it("throws when deleting a part-level relationship that does not exist", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    await expect(docPart.deleteRelationship("rIdDoesNotExist")).rejects.toThrow(
      "Relationship not found: rIdDoesNotExist",
    );
  });

  it("getXDocument materializes a lazy blob-opened part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const xDoc = await docPart.getXDocument();
    expect(xDoc).toBeDefined();
    expect(xDoc.root).not.toBeNull();
    expect(xDoc.root!.element(W.body)).not.toBeNull();
  });

  it("getXDocument returns the XDocument directly for a FlatOPC-opened part", async () => {
    const pkg = await OpenXmlPackage.open(blankDocumentFlatOpc);
    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const xDoc = await docPart.getXDocument();
    expect(xDoc).toBeDefined();
    expect(xDoc.root!.element(W.body)).not.toBeNull();
  });

  it("getXDocument throws for a non-xml part", async () => {
    const pkg = await OpenXmlPackage.open(blankDocumentBase64);
    const imgPart = pkg.addPart("/word/media/img.png", "image/png", "binary", "fakedata");
    await expect(imgPart.getXDocument()).rejects.toThrow(
      "Cannot get XDocument for non-xml part: /word/media/img.png",
    );
  });

  it("putXDocument throws when xDoc is null", async () => {
    const pkg = await OpenXmlPackage.open(blankDocumentBase64);
    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    expect(() => docPart.putXDocument(null as unknown as XDocument)).toThrow(
      "putXDocument: xDoc must not be null or undefined",
    );
  });

  it("getParts returns related parts of document part from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const parts = await docPart.getParts();

    expect(parts.length).toBeGreaterThan(0);
    expect(parts.some((p) => p.getUri() === "/word/comments.xml")).toBe(true);
  });

  it("getParts throws for a dangling internal relationship", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    await docPart.addRelationship("rId99", RelationshipType.wordprocessingComments, "missing.xml");

    await expect(docPart.getParts()).rejects.toThrow(
      "Part not found for relationship target: /word/missing.xml",
    );
  });
});
