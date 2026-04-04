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

/**
 * Excel spreadsheet package — opens, navigates, and saves `.xlsx` files.
 *
 * @remarks
 * Extends {@link OpenXmlPackage} with Excel-specific convenience methods.
 * All parts returned by this class are typed as {@link SmlPart}.
 *
 * @example
 * ```typescript
 * import { SmlPackage, S } from "openxmlsdkts";
 * import fs from "fs";
 *
 * const buffer = fs.readFileSync("data.xlsx");
 * const doc = await SmlPackage.open(new Blob([buffer]));
 * const workbook = await doc.workbookPart();
 * const worksheets = await workbook!.worksheetParts();
 * console.log(`${worksheets.length} worksheets`);
 * ```
 *
 * @category Class and Type Reference
 */
export class SmlPackage extends OpenXmlPackage {
  /**
   * Opens an Excel document from any supported format.
   *
   * @param document - The document to open (Blob, Base64 string, or Flat OPC XML string).
   * @returns A promise resolving to a {@link SmlPackage} instance.
   */
  static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<SmlPackage> {
    return OpenXmlPackage.openInto(new SmlPackage(), document);
  }

  /** @internal */
  protected createPart(pkg: OpenXmlPackage, uri: string, contentType: string | null, partType: PartType, data: unknown): SmlPart {
    return new SmlPart(pkg, uri, contentType, partType, data);
  }

  /**
   * Returns the main workbook part.
   *
   * @returns The workbook {@link SmlPart}, or `undefined` if not found.
   */
  async workbookPart(): Promise<SmlPart | undefined> {
    return (await this.getPartsByContentType(ContentType.workbook))[0] as SmlPart | undefined;
  }
}
