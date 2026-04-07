/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from "vitest";
import { WmlPackage, XElement, W } from "openxmlsdkts";

const inputXml = `\
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:pStyle w:val="ListParagraph"/>
    <w:spacing w:before="60" w:after="100" w:line="240" w:lineRule="auto"/>
    <w:contextualSpacing w:val="0"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:b/>
      <w:bCs/>
    </w:rPr>
    <w:t>123</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    </w:rPr>
    <w:t xml:space="preserve"> abc </w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:b/>
      <w:bCs/>
    </w:rPr>
    <w:t>456</w:t>
  </w:r>
</w:p>`;

const characterLevelExpected = `<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:pPr>
    <w:pStyle w:val='ListParagraph' />
    <w:spacing w:before='60' w:after='100' w:line='240' w:lineRule='auto' />
    <w:contextualSpacing w:val='0' />
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
      <w:b />
      <w:bCs />
    </w:rPr>
    <w:t>1</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
      <w:b />
      <w:bCs />
    </w:rPr>
    <w:t>2</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
      <w:b />
      <w:bCs />
    </w:rPr>
    <w:t>3</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
    </w:rPr>
    <w:t xml:space='preserve'> </w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
    </w:rPr>
    <w:t>a</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
    </w:rPr>
    <w:t>b</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
    </w:rPr>
    <w:t>c</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
    </w:rPr>
    <w:t xml:space='preserve'> </w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
      <w:b />
      <w:bCs />
    </w:rPr>
    <w:t>4</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
      <w:b />
      <w:bCs />
    </w:rPr>
    <w:t>5</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii='Times New Roman' w:eastAsia='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman' />
      <w:b />
      <w:bCs />
    </w:rPr>
    <w:t>6</w:t>
  </w:r>
</w:p>`;

const inputWithNonTextRuns = `\
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:pStyle w:val="Normal"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/>
    </w:rPr>
    <w:t>AB</w:t>
  </w:r>
  <w:r>
    <w:br/>
  </w:r>
  <w:r>
    <w:rPr>
      <w:b/>
    </w:rPr>
    <w:t>CD</w:t>
  </w:r>
</w:p>`;

describe("atomizeRunsInParagraph", () => {
  it("splits text runs into character-level runs", () => {
    const para = XElement.parse(inputXml);
    const result = WmlPackage.atomizeRunsInParagraph(para);
    expect(result.toStringWithIndentation()).toBe(characterLevelExpected);
  });

  it("preserves non-text runs", () => {
    const para = XElement.parse(inputWithNonTextRuns);
    const result = WmlPackage.atomizeRunsInParagraph(para);

    const runs = result.elements(W.r);
    // A, B, <br run>, C, D = 5 runs
    expect(runs).toHaveLength(5);

    // The break run (index 2) should have w:br and no w:t
    const brRun = runs[2];
    expect(brRun.element(W.br)).toBeDefined();
    expect(brRun.element(W.t)).toBeNull();

    // Text runs should each have one character
    expect(runs[0].element(W.t)!.value).toBe("A");
    expect(runs[1].element(W.t)!.value).toBe("B");
    expect(runs[3].element(W.t)!.value).toBe("C");
    expect(runs[4].element(W.t)!.value).toBe("D");
  });
});

describe("coalesceRunsInParagraph", () => {
  it("merges adjacent same-formatted runs", () => {
    const para = XElement.parse(inputXml);
    const atomized = WmlPackage.atomizeRunsInParagraph(para);
    const coalesced = WmlPackage.coalesceRunsInParagraph(atomized);

    expect(coalesced.toStringWithIndentation()).toBe(XElement.parse(inputXml).toStringWithIndentation());
  });

  it("preserves non-text runs and breaks adjacency", () => {
    const para = XElement.parse(inputWithNonTextRuns);
    const atomized = WmlPackage.atomizeRunsInParagraph(para);
    const coalesced = WmlPackage.coalesceRunsInParagraph(atomized);

    const runs = coalesced.elements(W.r);
    // "AB" (coalesced), <br run>, "CD" (coalesced) = 3 runs
    expect(runs).toHaveLength(3);
    expect(runs[0].element(W.t)!.value).toBe("AB");
    expect(runs[1].element(W.br)).toBeDefined();
    expect(runs[2].element(W.t)!.value).toBe("CD");
  });
});

describe("atomize/coalesce round-trip", () => {
  it("round-trips to the original paragraph", () => {
    const para = XElement.parse(inputXml);
    const atomized = WmlPackage.atomizeRunsInParagraph(para);
    expect(atomized.toStringWithIndentation()).toBe(characterLevelExpected);

    const coalesced = WmlPackage.coalesceRunsInParagraph(atomized);
    expect(coalesced.toStringWithIndentation()).toBe(XElement.parse(inputXml).toStringWithIndentation());
  });

  it("round-trips a paragraph with non-text runs", () => {
    const para = XElement.parse(inputWithNonTextRuns);
    const atomized = WmlPackage.atomizeRunsInParagraph(para);
    const coalesced = WmlPackage.coalesceRunsInParagraph(atomized);

    expect(coalesced.toStringWithIndentation()).toBe(XElement.parse(inputWithNonTextRuns).toStringWithIndentation());
  });
});
