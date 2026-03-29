/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { PmlPackage } from "OpenXmlSdkTs";
import * as fs from "fs";
import * as path from "path";

describe("PmlPackage", () => {
  it("does not throw when opening a pptx blob", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    await expect(PmlPackage.open(blob)).resolves.toBeDefined();
  });

  it("presentationPart returns the presentation part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const part = await doc.presentationPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/ppt/presentation.xml");
  });

  it("coreFilePropertiesPart returns the core properties part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const part = await doc.coreFilePropertiesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/docProps/core.xml");
  });

  it("extendedFilePropertiesPart returns the app properties part", async () => {
    const srcFile = path.resolve(__dirname, "../../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const part = await doc.extendedFilePropertiesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/docProps/app.xml");
  });
});
