/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { SmlPackage } from "openxmlsdkts";
import * as fs from "fs";
import * as path from "path";

describe("SmlPackage", () => {
  it("does not throw when opening an xlsx blob", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    await expect(SmlPackage.open(blob)).resolves.toBeDefined();
  });

  it("workbookPart returns the workbook part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await SmlPackage.open(blob);
    const part = await doc.workbookPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/xl/workbook.xml");
  });

  it("coreFilePropertiesPart returns the core properties part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await SmlPackage.open(blob);
    const part = await doc.coreFilePropertiesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/docProps/core.xml");
  });

  it("extendedFilePropertiesPart returns the app properties part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await SmlPackage.open(blob);
    const part = await doc.extendedFilePropertiesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/docProps/app.xml");
  });
});
