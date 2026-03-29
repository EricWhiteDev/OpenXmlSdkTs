/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { WmlDocument } from "OpenXmlSdkTs";
import * as fs from "fs";
import * as path from "path";

describe("WmlDocument", () => {
  it("does not throw when opening a docx blob", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    await expect(WmlDocument.open(blob)).resolves.toBeDefined();
  });

  it("mainDocumentPart returns the main document part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/TemplateDocument.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlDocument.open(blob);
    const part = await doc.mainDocumentPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/word/document.xml");
  });

  it("contentParts returns main document, headers, and footers", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/WithPageHeaderAndPageFooter.docx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await WmlDocument.open(blob);
    const parts = await doc.contentParts();
    const uris = parts.map((p) => p.getUri());
    expect(uris).toContain("/word/document.xml");
    expect(uris).toContain("/word/header1.xml");
    expect(uris).toContain("/word/footer1.xml");
    expect(parts[0].getUri()).toBe("/word/document.xml");
  });
});
