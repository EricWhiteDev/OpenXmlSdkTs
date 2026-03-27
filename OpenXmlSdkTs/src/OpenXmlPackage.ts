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

export class OpenXmlPackage {
    private parts: Map<string, OpenXmlPart> = new Map();
    private ctXDoc!: XDocument; // This is the XDocument for the content types in the package

    constructor(document: unknown) {
        if (typeof document === 'string') {
            if (OpenXmlUtility.isBase64(document)) {
                OpenXmlPackage.openFromBase64Internal(this, document);
            } else {
                OpenXmlPackage.openFromFlatOpcInternal(this, document);
            }
        } else {
            OpenXmlPackage.openFromBlobInternal(this, document);
        }
    }

    private static openFromBase64Internal(pkg: OpenXmlPackage, document: string): void {
    }

    private static openFromFlatOpcInternal(pkg: OpenXmlPackage, document: string): void {
    }

    private static openFromBlobInternal(pkg: OpenXmlPackage, document: unknown): void {
    }
}
