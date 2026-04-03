/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { PmlPackage, PmlPart } from "openxmlsdkts";
import * as fs from "fs";
import * as path from "path";

describe("PmlPart", () => {
  it("presentationPart returns a PmlPart instance", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const part = await doc.presentationPart();
    expect(part).toBeInstanceOf(PmlPart);
  });

  it("presentationPart slideParts returns all slide parts", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const pres = await doc.presentationPart();
    const slides = await pres!.slideParts();
    expect(slides.length).toBe(2);
    expect(slides.some((p) => p.getUri() === "/ppt/slides/slide1.xml")).toBe(true);
    expect(slides.some((p) => p.getUri() === "/ppt/slides/slide2.xml")).toBe(true);
  });

  it("presentationPart slideMasterParts returns slide master parts", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const pres = await doc.presentationPart();
    const masters = await pres!.slideMasterParts();
    expect(masters.length).toBeGreaterThanOrEqual(1);
    expect(masters.some((p) => p.getUri() === "/ppt/slideMasters/slideMaster1.xml")).toBe(true);
  });

  it("presentationPart viewPropertiesPart returns the view properties part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const pres = await doc.presentationPart();
    const part = await pres!.viewPropertiesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/ppt/viewProps.xml");
  });

  it("presentationPart presentationPropertiesPart returns the presentation properties part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const pres = await doc.presentationPart();
    const part = await pres!.presentationPropertiesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/ppt/presProps.xml");
  });

  it("presentationPart tableStylesPart returns the table styles part", async () => {
    const srcFile = path.resolve(__dirname, "../test-files/Sample.pptx");
    const buffer = fs.readFileSync(srcFile);
    const blob = new Blob([buffer]);
    const doc = await PmlPackage.open(blob);
    const pres = await doc.presentationPart();
    const part = await pres!.tableStylesPart();
    expect(part).toBeDefined();
    expect(part!.getUri()).toBe("/ppt/tableStyles.xml");
  });
});
