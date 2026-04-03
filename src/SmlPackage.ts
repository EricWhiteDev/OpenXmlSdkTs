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
import { SmlPart } from "./SmlPart";

export class SmlPackage extends OpenXmlPackage {
  static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<SmlPackage> {
    return OpenXmlPackage.openInto(new SmlPackage(), document);
  }

  protected createPart(pkg: OpenXmlPackage, uri: string, contentType: string | null, partType: PartType, data: unknown): SmlPart {
    return new SmlPart(pkg, uri, contentType, partType, data);
  }

  async workbookPart(): Promise<SmlPart | undefined> {
    return (await this.getPartsByContentType(ContentType.workbook))[0] as SmlPart | undefined;
  }
}
