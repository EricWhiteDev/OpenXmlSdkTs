/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect, vi } from "vitest";
import {
  OpenXmlPackage,
  W,
  ContentType,
  RelationshipType,
  XDocument,
  XElement,
  XAttribute,
} from "OpenXmlSdkTs";
import { blankDocumentBase64, blankDocumentFlatOpc } from "./TestResources";
import JSZip from "jszip";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";

describe("OpenXmlPackage", () => {
  it("does not throw when opening a docx blob", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const tmpFile = path.join(os.tmpdir(), `openxmlpackage-test-${Date.now()}.docx`);
    fs.copyFileSync(srcFile, tmpFile);
    try {
      const buffer = fs.readFileSync(tmpFile);
      const blob = new Blob([buffer]);
      const spy = vi.spyOn(JSZip, "loadAsync");
      await expect(OpenXmlPackage.open(blob)).resolves.toBeDefined();
      expect(spy).toHaveBeenCalledWith(expect.any(ArrayBuffer));
    } finally {
      fs.unlinkSync(tmpFile);
    }
  });

  it("opens a base64-encoded docx via openFromBase64Internal", async () => {
    const spy = vi.spyOn(JSZip, "loadAsync");
    await expect(OpenXmlPackage.open(blankDocumentBase64)).resolves.toBeDefined();
    expect(spy).toHaveBeenCalledWith(expect.any(String), { base64: true });
  });

  it("opens a FlatOPC string via openFlatOpcFromXDoc", async () => {
    const spy = vi.spyOn(JSZip, "loadAsync");
    await expect(OpenXmlPackage.open(blankDocumentFlatOpc)).resolves.toBeDefined();
    expect(spy).not.toHaveBeenCalled();
  });

  it("opens a docx blob with the correct parts", async () => {
    const expectedParts = [
      {
        uri: "/word/document.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      },
      {
        uri: "/word/theme/theme1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
      },
      {
        uri: "/word/settings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
      },
      {
        uri: "/word/fontTable.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
      },
      {
        uri: "/word/webSettings.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
      },
      {
        uri: "/docProps/app.xml",
        contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      },
      {
        uri: "/docProps/core.xml",
        contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      },
      {
        uri: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      },
    ];
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const actualParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));
    expect(actualParts).toEqual(expectedParts);
  });

  it("opens a base64-encoded docx with the correct parts", async () => {
    const expectedParts = [
      {
        uri: "/word/document.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      },
      {
        uri: "/word/theme/theme1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
      },
      {
        uri: "/word/settings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
      },
      {
        uri: "/word/fontTable.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
      },
      {
        uri: "/word/webSettings.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
      },
      {
        uri: "/docProps/app.xml",
        contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      },
      {
        uri: "/docProps/core.xml",
        contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      },
      {
        uri: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      },
    ];
    const pkg = await OpenXmlPackage.open(blankDocumentBase64);
    const actualParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));
    expect(actualParts).toEqual(expectedParts);
  });

  it("opens a FlatOPC string with the correct parts", async () => {
    const expectedParts = [
      {
        uri: "/word/document.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      },
      {
        uri: "/word/theme/theme1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
      },
      {
        uri: "/word/settings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
      },
      {
        uri: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      },
      {
        uri: "/word/webSettings.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
      },
      {
        uri: "/word/fontTable.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
      },
      {
        uri: "/docProps/core.xml",
        contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      },
      {
        uri: "/docProps/app.xml",
        contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      },
    ];
    const pkg = await OpenXmlPackage.open(blankDocumentFlatOpc);
    const actualParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));
    expect(actualParts).toEqual(expectedParts);
  });

  it("opens WithComments.docx with the correct parts", async () => {
    const expectedParts = [
      {
        uri: "/word/document.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      },
      {
        uri: "/word/comments.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
      },
      {
        uri: "/word/theme/theme1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
      },
      {
        uri: "/word/settings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
      },
      {
        uri: "/word/fontTable.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
      },
      {
        uri: "/word/webSettings.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
      },
      {
        uri: "/docProps/app.xml",
        contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      },
      {
        uri: "/docProps/core.xml",
        contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      },
      {
        uri: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      },
      {
        uri: "/word/people.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",
      },
      {
        uri: "/word/commentsExtended.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
      },
    ];
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const actualParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));
    expect(actualParts).toEqual(expectedParts);
  });

  it("saves a FlatOPC-opened package back to FlatOPC with the correct parts", async () => {
    const pkg = await OpenXmlPackage.open(blankDocumentFlatOpc);
    const originalParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    const saved = await pkg.saveToFlatOpcAsync();
    const pkg2 = await OpenXmlPackage.open(saved);
    const roundTrippedParts = pkg2
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    expect(roundTrippedParts).toEqual(originalParts);
  });

  it("saves a blob-opened package to FlatOPC with the correct parts", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const originalParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    const saved = await pkg.saveToFlatOpcAsync();
    const pkg2 = await OpenXmlPackage.open(saved);
    const roundTrippedParts = pkg2
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    expect(roundTrippedParts).toEqual(originalParts);
  });

  it("saves a FlatOPC-opened package to base64 and back", async () => {
    const pkg = await OpenXmlPackage.open(blankDocumentFlatOpc);
    const originalParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);
    const roundTrippedParts = pkg2
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    expect(roundTrippedParts).toEqual(originalParts);
  });

  it("saves a blob-opened package to base64 and back", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const originalParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);
    const roundTrippedParts = pkg2
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    expect(roundTrippedParts).toEqual(originalParts);
  });

  it("saves a blob-opened package to a Blob and back", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const originalParts = pkg
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    const saved = await pkg.saveToBlobAsync();
    const pkg2 = await OpenXmlPackage.open(saved);
    const roundTrippedParts = pkg2
      .getParts()
      .map((p) => ({ uri: p.getUri(), contentType: p.getContentType() }));

    expect(roundTrippedParts).toEqual(originalParts);
  });

  it("adds a comments part to a blank document and round-trips correctly", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/BlankDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const docXDoc = await docPart.getXDocument();

    const sectPr = docXDoc.root!.element(W.body)!.element(W.sectPr)!;
    const paragraph = new XElement(
      W.p,
      new XElement(W.commentRangeStart, new XAttribute(W.id, "0")),
      new XElement(W.r, new XElement(W.t, "Commented text")),
      new XElement(W.commentRangeEnd, new XAttribute(W.id, "0")),
      new XElement(
        W.r,
        new XElement(W.rPr, new XElement(W.rStyle, new XAttribute(W.val, "CommentReference"))),
        new XElement(W.commentReference, new XAttribute(W.id, "0")),
      ),
    );
    sectPr.addBeforeSelf(paragraph);

    const commentsXDoc = new XDocument(
      new XElement(
        W.comments,
        new XElement(
          W.comment,
          new XAttribute(W.id, "0"),
          new XAttribute(W.author, "Test Author"),
          new XAttribute(W.date, "2026-01-01T00:00:00Z"),
          new XElement(W.p, new XElement(W.r, new XElement(W.t, "A comment"))),
        ),
      ),
    );

    pkg.addPart("/word/comments.xml", ContentType.wordprocessingComments, "xml", commentsXDoc);

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);

    const commentsPart = pkg2.getParts().find((p) => p.getUri() === "/word/comments.xml");
    expect(commentsPart).toBeDefined();

    const commentsXDoc2 = await commentsPart!.getXDocument();
    const commentEl = commentsXDoc2.root!.element(W.comment);
    expect(commentEl).toBeDefined();
    expect(commentEl!.attribute(W.id)?.value).toBe("0");
  });

  it("gets package-level relationships from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const rels = await pkg.getRelationships();

    const mainDocRel = rels.find((r) => r.getType() === RelationshipType.mainDocument);
    expect(mainDocRel).toBeDefined();
    expect(mainDocRel!.getTarget()).toBe("word/document.xml");
  });

  it("gets package-level relationships by type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const rels = await pkg.getRelationshipsByRelationshipType(RelationshipType.mainDocument);

    expect(rels).toHaveLength(1);
    expect(rels[0].getTarget()).toBe("word/document.xml");
  });

  it("gets parts by relationship type from package level in WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const parts = await pkg.getPartsByRelationshipType(RelationshipType.mainDocument);

    expect(parts).toHaveLength(1);
    expect(parts[0].getUri()).toBe("/word/document.xml");
  });

  it("gets part by relationship type from package level in WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const part = await pkg.getPartByRelationshipType(RelationshipType.mainDocument);

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/document.xml");
  });

  it("adds a package-level relationship and round-trips correctly", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const rel = await pkg.addRelationship(
      "rId99",
      RelationshipType.extendedFileProperties,
      "docProps/custom.xml",
    );
    expect(rel.getId()).toBe("rId99");
    expect(rel.getType()).toBe(RelationshipType.extendedFileProperties);
    expect(rel.getTarget()).toBe("docProps/custom.xml");
    expect(rel.getTargetMode()).toBeNull();

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);
    const roundTrippedRel = await pkg2.getRelationshipById("rId99");
    expect(roundTrippedRel).toBeDefined();
    expect(roundTrippedRel!.getType()).toBe(RelationshipType.extendedFileProperties);
    expect(roundTrippedRel!.getTarget()).toBe("docProps/custom.xml");
  });

  it("gets a package-level relationship by id from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const rel = await pkg.getRelationshipById("rId1");

    expect(rel).toBeDefined();
    expect(rel!.getType()).toBe(RelationshipType.mainDocument);
    expect(rel!.getTarget()).toBe("word/document.xml");
  });

  it("gets a package-level part by id from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const part = await pkg.getPartById("rId1");

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/document.xml");
  });

  it("gets package-level relationships by content type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const rels = await pkg.getRelationshipsByContentType(ContentType.mainDocument);

    expect(rels).toHaveLength(1);
    expect(rels[0].getTarget()).toBe("word/document.xml");
  });

  it("gets package-level parts by content type from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const parts = await pkg.getPartsByContentType(ContentType.mainDocument);

    expect(parts).toHaveLength(1);
    expect(parts[0].getUri()).toBe("/word/document.xml");
  });

  it("saves a FlatOPC-opened package to FlatOPC preserving document body content", async () => {
    const pkg = await OpenXmlPackage.open(blankDocumentFlatOpc);
    const saved = await pkg.saveToFlatOpcAsync();
    const pkg2 = await OpenXmlPackage.open(saved);
    const docPart = pkg2.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const docXDoc = await docPart.getXDocument();
    const body = docXDoc.root!.element(W.body);
    expect(body).not.toBeNull();
    expect(body!.element(W.sectPr)).not.toBeNull();
  });

  it("saves a blob-opened package to FlatOPC preserving document body content", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const saved = await pkg.saveToFlatOpcAsync();
    const pkg2 = await OpenXmlPackage.open(saved);
    const docPart = pkg2.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const docXDoc = await docPart.getXDocument();
    const body = docXDoc.root!.element(W.body);
    expect(body).not.toBeNull();
    expect(body!.element(W.sectPr)).not.toBeNull();
  });

  it("contentParts returns main document, headers, and footers", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithPageHeaderAndPageFooter.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const parts = await pkg.contentParts();
    const uris = parts.map((p) => p.getUri());
    expect(uris).toContain("/word/document.xml");
    expect(uris).toContain("/word/header1.xml");
    expect(uris).toContain("/word/footer1.xml");
    expect(parts[0].getUri()).toBe("/word/document.xml");
  });

  it("mainDocumentPart returns the main document part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const part = await pkg.mainDocumentPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/document.xml");
  });

  it("coreFilePropertiesPart returns the core properties part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const part = await pkg.coreFilePropertiesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/docProps/core.xml");
  });

  it("extendedFilePropertiesPart returns the app properties part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    const part = await pkg.extendedFilePropertiesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/docProps/app.xml");
  });

  it("workbookPart returns undefined for a Word document", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    expect(await pkg.workbookPart()).toBeUndefined();
  });

  it("customFilePropertiesPart returns undefined when no custom properties exist", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    expect(await pkg.customFilePropertiesPart()).toBeUndefined();
  });

  it("presentationPart returns undefined for a Word document", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    expect(await pkg.presentationPart()).toBeUndefined();
  });

  it("returns the content type for a known URI", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    expect(pkg.getContentType("/word/document.xml")).toBe(ContentType.mainDocument);
  });

  it("throws when getting the content type for an unknown URI", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);
    expect(() => pkg.getContentType("/word/doesNotExist.xyz")).toThrow(
      "Content type not found for part: /word/doesNotExist.xyz",
    );
  });

  it("deletes a package-level relationship and verifies it is gone after round-trip", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const result = await pkg.deleteRelationship("rId1");
    expect(result).toBe(true);
    expect(await pkg.getRelationshipById("rId1")).toBeUndefined();

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);
    expect(await pkg2.getRelationshipById("rId1")).toBeUndefined();
  });

  it("throws when deleting a package-level relationship that does not exist", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    await expect(pkg.deleteRelationship("rIdDoesNotExist")).rejects.toThrow(
      "Relationship not found: rIdDoesNotExist",
    );
  });

  it("deletes the comments part from a document with comments and round-trips correctly", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const docXDoc = await docPart.getXDocument();

    for (const el of docXDoc.root!.descendants(W.commentRangeStart)) {
      el.remove();
    }
    for (const el of docXDoc.root!.descendants(W.commentRangeEnd)) {
      el.remove();
    }
    for (const el of docXDoc.root!.descendants(W.r)) {
      if (el.element(W.commentReference) !== null) {
        el.remove();
      }
    }
    const commentsPart = pkg.getParts().find((p) => p.getUri() === "/word/comments.xml")!;
    await pkg.deletePart(commentsPart);

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);

    expect(pkg2.getParts().find((p) => p.getUri() === "/word/comments.xml")).toBeUndefined();

    const docPart2 = pkg2.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const docXDoc2 = await docPart2.getXDocument();
    expect(docXDoc2.root!.descendants(W.commentRangeStart).length).toBe(0);
  });
});
