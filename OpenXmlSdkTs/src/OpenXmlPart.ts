/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { OpenXmlPackage } from "./OpenXmlPackage";

export type PartType = "binary" | "base64" | "xml" | null;

export class OpenXmlPart {
  private pkg: OpenXmlPackage;
  private uri: string;
  private contentType: string | null; // MIME type value, e.g. "application/vnd...+xml"
  private partType: PartType;
  private data: unknown; // for now, this type is unknown.  May change later.

  constructor(
    pkg: OpenXmlPackage,
    uri: string,
    contentType: string | null,
    partType: PartType,
    data: unknown,
  ) {
    this.pkg = pkg;
    this.uri = uri;
    this.contentType = contentType;
    this.partType = partType;
    this.data = data;
  }

  getUri(): string {
    return this.uri;
  }

  getContentType(): string | null {
    return this.contentType;
  }

  getData(): unknown {
    return this.data;
  }

  setContentType(ct: string): void {
    this.contentType = ct;
  }

  getPartType(): PartType {
    return this.partType;
  }

  setData(data: unknown): void {
    this.data = data;
  }

  setPartType(pt: PartType): void {
    this.partType = pt;
  }
}
