/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { SmlPackage, SmlPart } from "openxmlsdkts";
import * as fs from "fs";
import * as path from "path";

describe("SmlPart", () => {
  it("workbookPart returns an SmlPart instance", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await SmlPackage.open(blob);
    const part = await doc.workbookPart();
    expect(part).toBeInstanceOf(SmlPart);
  });

  it("workbookPart worksheetParts returns all worksheet parts", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await SmlPackage.open(blob);
    const wb = await doc.workbookPart();
    const sheets = await wb!.worksheetParts();
    expect(sheets.length).toBeGreaterThanOrEqual(2);
    expect(sheets.some((p) => p.getUri() === "/xl/worksheets/sheet1.xml")).toBe(true);
    expect(sheets.some((p) => p.getUri() === "/xl/worksheets/sheet2.xml")).toBe(true);
  });

  it("workbookPart workbookStylesPart returns the styles part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await SmlPackage.open(blob);
    const wb = await doc.workbookPart();
    const part = await wb!.workbookStylesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/xl/styles.xml");
  });

  it("workbookPart sharedStringTablePart returns the shared strings part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await SmlPackage.open(blob);
    const wb = await doc.workbookPart();
    const part = await wb!.sharedStringTablePart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/xl/sharedStrings.xml");
  });

  it("workbookPart calculationChainPart returns the calculation chain part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.xlsx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await SmlPackage.open(blob);
    const wb = await doc.workbookPart();
    const part = await wb!.calculationChainPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/xl/calcChain.xml");
  });
});
