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

export class PmlPackage extends OpenXmlPackage {
  static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<PmlPackage> {
    return OpenXmlPackage.openInto(new PmlPackage(), document);
  }

  async presentationPart(): Promise<OpenXmlPart | undefined> {
    return (await this.getPartsByContentType(ContentType.presentation))[0];
  }
}
