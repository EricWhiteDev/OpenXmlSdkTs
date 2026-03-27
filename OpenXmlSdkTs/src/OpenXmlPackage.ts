/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { XDocument } from 'ltxmlts';
import { OpenXmlPart } from './OpenXmlPart';
import { OpenXmlUtility } from './OpenXmlUtility';

export type Base64String = string;
export type FlatOpcString = string;
export type DocxBinary = Blob;

export class OpenXmlPackage {
    private parts: Map<string, OpenXmlPart> = new Map();
    private ctXDoc!: XDocument; // This is the XDocument for the content types in the package

    constructor(document: Base64String | FlatOpcString | DocxBinary) {
        if (typeof document === 'string') {
            if (OpenXmlUtility.isBase64(document)) {
                OpenXmlPackage.openFromBase64Internal(this, document);
            } else {
                OpenXmlPackage.openFromFlatOpcInternal(this, document);
            }
        } else if (document instanceof Blob) {
            OpenXmlPackage.openFromBlobInternal(this, document);
        } else {
            throw new Error('Invalid argument: document must be a Base64String, FlatOpcString, or DocxBinary (Blob).');
        }
    }

    private static openFromBase64Internal(pkg: OpenXmlPackage, document: Base64String): void {
    }

    private static openFromFlatOpcInternal(pkg: OpenXmlPackage, document: FlatOpcString): void {
    }

    private static openFromBlobInternal(pkg: OpenXmlPackage, document: DocxBinary): void {
    }
}
