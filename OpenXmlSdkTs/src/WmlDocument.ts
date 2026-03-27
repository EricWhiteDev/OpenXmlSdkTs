/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import * as fs from 'fs';
import { XDocument } from 'ltxmlts';

export class WmlDocument {
  private constructor(
    private readonly filePath: string | null,
    private readonly xdocument: XDocument | null
  ) {}

  static open(fileName: string): WmlDocument;
  static open(xdocument: XDocument): WmlDocument;
  static open(arg: string | XDocument): WmlDocument {
    if (typeof arg === 'string') {
      if (!fs.existsSync(arg)) {
        throw new Error(`File not found: ${arg}`);
      }
      return new WmlDocument(arg, null);
    } else {
      if (arg.root === null) {
        throw new Error('XDocument has no root element');
      }
      return new WmlDocument(null, arg);
    }
  }
}
