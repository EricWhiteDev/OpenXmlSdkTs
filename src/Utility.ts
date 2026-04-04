/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import type { OpenXmlPart } from "./OpenXmlPart";

/**
 * Static helper utilities for working with Open XML packages.
 *
 * @example
 * ```typescript
 * import { Utility } from "openxmlsdkts";
 *
 * const relsUri = Utility.getRelsPartUri(mainPart); // "/word/_rels/document.xml.rels"
 * Utility.isBase64("UEsDBBQAAAA..."); // true
 * Utility.isBase64("<?xml ...");      // false
 * ```
 */
export class Utility {
  /**
   * Computes the URI of the `.rels` file for a given part.
   *
   * @param part - The part whose `.rels` URI to compute.
   * @returns The `.rels` URI string (e.g., `"/word/_rels/document.xml.rels"`).
   */
  static getRelsPartUri(part: OpenXmlPart): string {
    const uri = part.getUri();
    const lastSlash = uri.lastIndexOf("/");
    return uri.substring(0, lastSlash) + "/_rels/" + uri.substring(lastSlash + 1) + ".rels";
  }

  /**
   * Returns the `.rels` part associated with a given part.
   *
   * @param part - The part whose `.rels` part to retrieve.
   * @returns The `.rels` {@link OpenXmlPart}, or `undefined` if none exists.
   */
  static getRelsPart(part: OpenXmlPart): OpenXmlPart | undefined {
    return part.getPkg().getPartByUri(Utility.getRelsPartUri(part));
  }

  /**
   * Heuristically determines whether a string is Base64-encoded.
   *
   * @remarks
   * Checks the first 500 characters for valid Base64 characters (`A-Z`, `a-z`, `0-9`, `+`, `/`).
   * Used by {@link OpenXmlPackage.open} to distinguish Base64 strings from Flat OPC XML.
   *
   * @param str - The value to test.
   * @returns `true` if the string appears to be Base64-encoded.
   */
  static isBase64(str: unknown): boolean {
    if (typeof str !== "string") {
      return false;
    }
    const sub = str.substring(0, 500);
    for (let i = 0; i < sub.length; i++) {
      const s = sub[i];
      if (s >= "A" && s <= "Z") {
        continue;
      }
      if (s >= "a" && s <= "z") {
        continue;
      }
      if (s >= "0" && s <= "9") {
        continue;
      }
      if (s === "+" || s === "/") {
        continue;
      }
      return false;
    }
    return true;
  }
}
