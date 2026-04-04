/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { ContentType } from "./ContentType";
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
