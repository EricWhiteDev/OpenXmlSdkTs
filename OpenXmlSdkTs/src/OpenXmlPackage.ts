/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { XDocument, XDeclaration, XElement, XAttribute } from 'ltxmlts';
import JSZip from 'jszip';
import { OpenXmlPart } from './OpenXmlPart';
import { OpenXmlUtility } from './OpenXmlUtility';
import { CT, FLATOPC } from './OpenXmlNamespacesAndNames';
import { ContentType } from './ContentType';

export type Base64String = string;
export type FlatOpcString = string;
export type DocxBinary = Blob;

export class OpenXmlPackage {
    private parts: Map<string, OpenXmlPart> = new Map();
    private ctXDoc!: XDocument; // This is the XDocument for the content types in the package

    static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<OpenXmlPackage> {
        const pkg = new OpenXmlPackage();
        if (typeof document === 'string') {
            if (OpenXmlUtility.isBase64(document)) {
                await OpenXmlPackage.openFromBase64Internal(pkg, document);
            } else {
                await OpenXmlPackage.openFromFlatOpcInternal(pkg, document);
            }
        } else if (document instanceof Blob) {
            await OpenXmlPackage.openFromBlobInternal(pkg, document);
        } else {
            throw new Error('Invalid argument: document must be a Base64String, FlatOpcString, or DocxBinary (Blob).');
        }
        return pkg;
    }

    private static getContentType(uri: string, ctXDoc: XDocument): string {
        const root = ctXDoc.root!;

        const override = root.elements(CT.Override)
            .find(el => el.attribute('PartName')?.value === uri);
        if (override) {
            const ct = override.attribute('ContentType')?.value;
            if (ct) return ct;
        }

        const ext = uri.split('.').pop() ?? '';
        const def = root.elements(CT.Default)
            .find(el => el.attribute('Extension')?.value === ext);
        if (def) {
            const ct = def.attribute('ContentType')?.value;
            if (ct) return ct;
        }

        throw new Error(`Content type not found for part: ${uri}`);
    }

    private static async openFromBase64Internal(pkg: OpenXmlPackage, document: Base64String): Promise<void> {
        const zip = await JSZip.loadAsync(document, { base64: true });
        await OpenXmlPackage.openFromZip(zip, pkg);
    }

    private static openFlatOpcFromXDoc(pkg: OpenXmlPackage, doc: XDocument): void {
        const root = doc.root!;
        pkg.ctXDoc = new XDocument(
            new XDeclaration("1.0", "utf-8", "yes"),
            new XElement(CT.Types,
                new XAttribute("xmlns", CT.namespace.namespaceName),
                new XElement(CT.Default,
                    new XAttribute("Extension", "rels"),
                    new XAttribute("ContentType", ContentType.relationships)),
                new XElement(CT.Default,
                    new XAttribute("Extension", "xml"),
                    new XAttribute("ContentType", "application/xml"))));

        for (const p of root.elements(FLATOPC.part)) {
            const uri = p.attribute(FLATOPC._name)!.value;
            const contentType = p.attribute(FLATOPC.contentType)!.value;
            const partType = contentType.endsWith('xml') ? 'xml' : 'base64';

            if (partType === 'xml') {
                const xmlDataEl = p.element(FLATOPC.xmlData)!;
                const newPart = new OpenXmlPart(pkg, uri, contentType, 'xml',
                    new XDocument(xmlDataEl.elements()[0]));
                pkg.parts.set(uri, newPart);
                if (contentType !== ContentType.relationships) {
                    pkg.ctXDoc.root!.add(
                        new XElement(CT.Override,
                            new XAttribute("PartName", uri),
                            new XAttribute("ContentType", contentType)));
                }
            } else {
                const binaryData = p.element(FLATOPC.binaryData)!.value;
                const newPart = new OpenXmlPart(pkg, uri, contentType, 'binary', binaryData);
                pkg.parts.set(uri, newPart);
                pkg.ctXDoc.root!.add(
                    new XElement(CT.Override,
                        new XAttribute("PartName", uri),
                        new XAttribute("ContentType", contentType)));
            }
        }

        const ctPart = new OpenXmlPart(pkg, "[Content_Types].xml", null, 'xml', pkg.ctXDoc);
        pkg.parts.set("[Content_Types].xml", ctPart);
    }

    private static async openFromFlatOpcInternal(pkg: OpenXmlPackage, document: FlatOpcString): Promise<void> {
        const xDoc = XDocument.parse(document);
        OpenXmlPackage.openFlatOpcFromXDoc(pkg, xDoc);
    }

    private static async openFromBlobInternal(pkg: OpenXmlPackage, document: DocxBinary): Promise<void> {
        const arrayBuffer = await document.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);
        await OpenXmlPackage.openFromZip(zip, pkg);
    }

    private static async openFromZip(zip: JSZip, pkg: OpenXmlPackage): Promise<void> {
        const ctZipFile = zip.files['[Content_Types].xml'];
        if (!ctZipFile)
            throw new Error('Invalid Open XML document: no [Content_Types].xml');
        const ctData = await ctZipFile.async('string');
        pkg.ctXDoc = XDocument.parse(ctData);

        for (const f in zip.files) {
            const zipFile = zip.files[f];
            if (!f.endsWith('/') && f !== '[Content_Types].xml') {
                const f2 = '/' + f;
                const newPart = new OpenXmlPart(pkg, f2, null, null, zipFile);
                pkg.parts.set(f2, newPart);
            }
        }

        for (const [part, thisPart] of pkg.parts) {
            const ct = OpenXmlPackage.getContentType(part, pkg.ctXDoc);
            thisPart.setContentType(ct);
            thisPart.setPartType(ct.endsWith('xml') ? 'xml' : 'binary');
        }
    }
}
