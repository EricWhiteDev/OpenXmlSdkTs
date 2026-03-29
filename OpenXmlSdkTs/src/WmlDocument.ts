/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { ContentType } from "./ContentType";
import { OpenXmlPackage, Base64String, FlatOpcString, DocxBinary } from "./OpenXmlPackage";
import { OpenXmlPart } from "./OpenXmlPart";

export class WmlDocument extends OpenXmlPackage {
  static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<WmlDocument> {
    return OpenXmlPackage.openInto(new WmlDocument(), document);
  }

  async mainDocumentPart(): Promise<OpenXmlPart | undefined> {
    return (await this.getPartsByContentType(ContentType.mainDocument))[0];
  }

  async contentParts(): Promise<OpenXmlPart[]> {
    const main = await this.mainDocumentPart();
    if (!main) {
      return [];
    }
    const parts: OpenXmlPart[] = [main];
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
