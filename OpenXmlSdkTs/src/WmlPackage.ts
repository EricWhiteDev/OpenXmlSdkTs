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
import { PartType } from "./OpenXmlPart";
import { WmlPart } from "./WmlPart";

export class WmlPackage extends OpenXmlPackage {
  static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<WmlPackage> {
    return OpenXmlPackage.openInto(new WmlPackage(), document);
  }

  protected createPart(pkg: OpenXmlPackage, uri: string, contentType: string | null, partType: PartType, data: unknown): WmlPart {
    return new WmlPart(pkg, uri, contentType, partType, data);
  }

  async mainDocumentPart(): Promise<WmlPart | undefined> {
    return (await this.getPartsByContentType(ContentType.mainDocument))[0] as WmlPart | undefined;
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
