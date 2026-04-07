/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { XElement, XAttribute, XNamespace, xseq } from "ltxmlts";
import { ContentType } from "./ContentType";
import { W } from "./OpenXmlNamespacesAndNames";
import { OpenXmlPackage, Base64String, FlatOpcString, OpcBinary } from "./OpenXmlPackage";
import { PartType } from "./OpenXmlPart";
import { WmlPart } from "./WmlPart";

/**
 * Word document package — opens, navigates, and saves `.docx` files.
 *
 * @remarks
 * Extends {@link OpenXmlPackage} with Word-specific convenience methods.
 * All parts returned by this class are typed as {@link WmlPart}, giving access
 * to Word-specific navigation like {@link WmlPart.headerParts | headerParts()},
 * {@link WmlPart.styleDefinitionsPart | styleDefinitionsPart()}, and more.
 *
 * @example
 * ```typescript
 * import { WmlPackage, W, XElement } from "openxmlsdkts";
 * import fs from "fs";
 *
 * const buffer = fs.readFileSync("report.docx");
 * const doc = await WmlPackage.open(new Blob([buffer]));
 * const mainPart = await doc.mainDocumentPart();
 * const xDoc = await mainPart!.getXDocument();
 * const body = xDoc.root!.element(W.body);
 * console.log(`Document has ${body!.elements(W.p).length} paragraphs`);
 *
 * // Add a new paragraph
 * const newPara = new XElement(W.p,
 *   new XElement(W.r, new XElement(W.t, "Hello from OpenXmlSdkTs!")));
 * body!.add(newPara);
 * mainPart!.putXDocument(xDoc);
 *
 * const blob = await doc.saveToBlobAsync();
 * ```
 *
 * @category Class and Type Reference
 */
export class WmlPackage extends OpenXmlPackage {
  /**
   * Opens a Word document from any supported format.
   *
   * @param document - The document to open (Blob, Base64 string, or Flat OPC XML string).
   * @returns A promise resolving to a {@link WmlPackage} instance.
   */
  static async open(document: Base64String | FlatOpcString | OpcBinary): Promise<WmlPackage> {
    return OpenXmlPackage.openInto(new WmlPackage(), document);
  }

  /** @internal */
  protected createPart(pkg: OpenXmlPackage, uri: string, contentType: string | null, partType: PartType, data: unknown): WmlPart {
    return new WmlPart(pkg, uri, contentType, partType, data);
  }

  /**
   * Returns the main document part (`/word/document.xml`).
   *
   * @returns The main document {@link WmlPart}, or `undefined` if not found.
   */
  async mainDocumentPart(): Promise<WmlPart | undefined> {
    return (await this.getPartsByContentType(ContentType.mainDocument))[0] as WmlPart | undefined;
  }

  /**
   * Returns all content-bearing parts: the main document, headers, footers, endnotes, and footnotes.
   *
   * @remarks
   * Useful for operations that need to process text across all content areas of the document.
   *
   * @returns An array of {@link WmlPart} instances.
   */
  /**
   * Splits each text run in a paragraph into one run per character.
   *
   * @remarks
   * Non-text runs (e.g., containing `w:br`, `w:tab`, `w:drawing`) and non-run
   * elements (e.g., `w:bookmarkStart`) are preserved in place as deep copies.
   * Paragraph properties (`w:pPr`) are cloned. Space characters get
   * `xml:space="preserve"` on their `w:t` element.
   *
   * @param para - A `w:p` element to atomize.
   * @returns A new `w:p` element with one character per text run.
   */
  static atomizeRunsInParagraph(para: XElement): XElement {
    const pPr = para.element(W.pPr);
    const newChildren: XElement[] = [];

    for (const child of para.elements()) {
      if (child.name.equals(W.pPr)) continue;

      if (child.name.equals(W.r) && child.element(W.t)) {
        const rPr = child.element(W.rPr);
        const text = child.element(W.t)!.value;
        for (const ch of [...text]) {
          const tElement = ch === " " ? new XElement(W.t, new XAttribute(XNamespace.xml + "space", "preserve"), ch) : new XElement(W.t, ch);
          newChildren.push(new XElement(W.r, rPr ? new XElement(rPr) : undefined, tElement));
        }
      } else {
        newChildren.push(new XElement(child));
      }
    }

    return new XElement(W.p, new XAttribute(XNamespace.xmlns + "w", W.namespace.namespaceName), pPr ? new XElement(pPr) : undefined, ...newChildren);
  }

  /**
   * Merges adjacent text runs that share the same run properties.
   *
   * @remarks
   * Uses `groupAdjacent` to find consecutive `w:r` elements whose `w:rPr`
   * serializes identically, then concatenates their `w:t` text into a single
   * run. Non-text runs and non-run elements break adjacency and are preserved
   * as deep copies. If the merged text starts or ends with a space,
   * `xml:space="preserve"` is added to the `w:t` element.
   *
   * @param para - A `w:p` element to coalesce.
   * @returns A new `w:p` element with adjacent same-formatted runs merged.
   */
  static coalesceRunsInParagraph(para: XElement): XElement {
    const pPr = para.element(W.pPr);

    const children = para.elements().filter((e) => !e.name.equals(W.pPr));
    let nonTextCounter = 0;
    const groups = xseq(children).groupAdjacent((el) => {
      if (el.name.equals(W.r) && el.element(W.t)) {
        return el.element(W.rPr)?.toString() ?? "";
      }
      return `\0${nonTextCounter++}`;
    });

    const newChildren: XElement[] = [];
    for (const group of groups) {
      const items = group.items.toArray() as XElement[];
      const first = items[0];

      if (first.name.equals(W.r) && first.element(W.t)) {
        const rPr = first.element(W.rPr);
        const text = items.map((r) => r.element(W.t)!.value).join("");
        const tElement =
          text.startsWith(" ") || text.endsWith(" ") ? new XElement(W.t, new XAttribute(XNamespace.xml + "space", "preserve"), text) : new XElement(W.t, text);
        newChildren.push(new XElement(W.r, rPr ? new XElement(rPr) : undefined, tElement));
      } else {
        for (const item of items) {
          newChildren.push(new XElement(item));
        }
      }
    }

    return new XElement(W.p, new XAttribute(XNamespace.xmlns + "w", W.namespace.namespaceName), pPr ? new XElement(pPr) : undefined, ...newChildren);
  }

  async contentParts(): Promise<WmlPart[]> {
    const main = await this.mainDocumentPart();
    if (!main) {
      return [];
    }
    const parts: WmlPart[] = [main];
    parts.push(...(await main.headerParts()));
    parts.push(...(await main.footerParts()));
    const endnotes = await main.endnotesPart();
    if (endnotes) {
      parts.push(endnotes);
    }
    const footnotes = await main.footnotesPart();
    if (footnotes) {
      parts.push(footnotes);
    }
    return parts;
  }
}
