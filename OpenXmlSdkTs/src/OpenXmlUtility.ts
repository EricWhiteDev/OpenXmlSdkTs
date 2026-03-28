/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

export class OpenXmlUtility {
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
