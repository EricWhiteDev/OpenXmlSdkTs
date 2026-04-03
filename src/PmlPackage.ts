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
import { PmlPart } from "./PmlPart";

export class PmlPackage extends OpenXmlPackage {
  static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<PmlPackage> {
    return OpenXmlPackage.openInto(new PmlPackage(), document);
  }

  protected createPart(pkg: OpenXmlPackage, uri: string, contentType: string | null, partType: PartType, data: unknown): PmlPart {
    return new PmlPart(pkg, uri, contentType, partType, data);
  }

  async presentationPart(): Promise<PmlPart | undefined> {
    return (await this.getPartsByContentType(ContentType.presentation))[0] as PmlPart | undefined;
  }
}
