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
    const docData = docPart.getData() as { async(type: string): Promise<string> };
    const docXDoc = XDocument.parse(await docData.async("string"));
    docPart.setData(docXDoc);

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

    const commentsData = commentsPart!.getData() as { async(type: string): Promise<string> };
    const commentsXDoc2 = XDocument.parse(await commentsData.async("string"));
    const commentEl = commentsXDoc2.root!.element(W.comment);
    expect(commentEl).toBeDefined();
    expect(commentEl!.attribute(W.id)?.value).toBe("0");
  });

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

  it("deletes the comments part from a document with comments and round-trips correctly", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const pkg = await OpenXmlPackage.open(blob);

    const docPart = pkg.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const docData = docPart.getData() as { async(type: string): Promise<string> };
    const docXDoc = XDocument.parse(await docData.async("string"));

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
    docPart.setData(docXDoc);

    const commentsPart = pkg.getParts().find((p) => p.getUri() === "/word/comments.xml")!;
    await pkg.deletePart(commentsPart);

    const saved = await pkg.saveToBase64Async();
    const pkg2 = await OpenXmlPackage.open(saved);

    expect(pkg2.getParts().find((p) => p.getUri() === "/word/comments.xml")).toBeUndefined();

    const docPart2 = pkg2.getParts().find((p) => p.getUri() === "/word/document.xml")!;
    const docData2 = docPart2.getData() as { async(type: string): Promise<string> };
    const docXDoc2 = XDocument.parse(await docData2.async("string"));
    expect(docXDoc2.root!.descendants(W.commentRangeStart).length).toBe(0);
  });
});
