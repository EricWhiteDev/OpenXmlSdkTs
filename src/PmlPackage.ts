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
import { PmlPart } from "./PmlPart";

/**
 * PowerPoint presentation package — opens, navigates, and saves `.pptx` files.
 *
 * @remarks
 * Extends {@link OpenXmlPackage} with PowerPoint-specific convenience methods.
 * All parts returned by this class are typed as {@link PmlPart}.
 *
 * @example
 * ```typescript
 * import { PmlPackage, P } from "openxmlsdkts";
 * import fs from "fs";
 *
 * const buffer = fs.readFileSync("deck.pptx");
 * const doc = await PmlPackage.open(new Blob([buffer]));
 * const presentation = await doc.presentationPart();
 * const slides = await presentation!.slideParts();
 * console.log(`${slides.length} slides`);
 * ```
 *
 * @category Class and Type Reference
 */
export class PmlPackage extends OpenXmlPackage {
  /**
   * Opens a PowerPoint document from any supported format.
   *
   * @param document - The document to open (Blob, Base64 string, or Flat OPC XML string).
   * @returns A promise resolving to a {@link PmlPackage} instance.
   */
  static async open(document: Base64String | FlatOpcString | OpcBinary): Promise<PmlPackage> {
    return OpenXmlPackage.openInto(new PmlPackage(), document);
  }

  /** @internal */
  protected createPart(pkg: OpenXmlPackage, uri: string, contentType: string | null, partType: PartType, data: unknown): PmlPart {
    return new PmlPart(pkg, uri, contentType, partType, data);
  }

  /**
   * Returns the main presentation part.
   *
   * @returns The presentation {@link PmlPart}, or `undefined` if not found.
   */
  async presentationPart(): Promise<PmlPart | undefined> {
    return (await this.getPartsByContentType(ContentType.presentation))[0] as PmlPart | undefined;
  }
}
