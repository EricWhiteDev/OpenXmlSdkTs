/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { WmlPackage, WmlPart } from "openxmlsdkts";
import * as fs from "fs";
import * as path from "path";

describe("WmlPart", () => {
  it("mainDocumentPart returns a WmlPart instance", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = await doc.mainDocumentPart();
    expect(docPart).toBeInstanceOf(WmlPart);
  });

  it("wordprocessingCommentsPart returns comments part from WithComments.docx", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/WithComments.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.wordprocessingCommentsPart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/comments.xml");
  });

  it("styleDefinitionsPart returns styles part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.styleDefinitionsPart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/styles.xml");
  });

  it("documentSettingsPart returns settings part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.documentSettingsPart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/settings.xml");
  });

  it("fontTablePart returns fontTable part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.fontTablePart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/fontTable.xml");
  });

  it("webSettingsPart returns webSettings part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.webSettingsPart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/webSettings.xml");
  });

  it("themePart returns theme part from TemplateDocument.docx", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const part = await docPart.themePart();

    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/theme/theme1.xml");
  });

  it("imageParts returns empty array on a Word document with no images", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlPackage.open(blob);
    const docPart = (await doc.mainDocumentPart())!;
    const parts = await docPart.imageParts();

    expect(parts).toHaveLength(0);
  });
});
