/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { XDocument } from 'ltxmlts';
import { ContentTypeKey } from './ContentType';
import { OpenXmlPackage } from './OpenXmlPackage';

export type PartType =
  | 'binary'
  | 'base64'
  | 'xml'
  | null;

export class OpenXmlPart {
  private pkg!: OpenXmlPackage;         // this is a reference to the parent OpenXmlPackage
  private uri!: string;                 // this is the uri of the part
  private contentType!: ContentTypeKey; // this is the content type of the part
  private partType!: PartType;
  private data: unknown;                // for now, this type is unknown.  May change later.
}
