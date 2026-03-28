/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect, vi } from 'vitest';
import { OpenXmlPackage } from 'OpenXmlSdkTs';
import { blankDocumentBase64, blankDocumentFlatOpc } from './TestResources';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

describe('OpenXmlPackage', () => {
    it('does not throw when opening a docx blob', async () => {
        const srcFile = path.resolve(__dirname, '../../test-files/TemplateDocument.docx');
        const tmpFile = path.join(os.tmpdir(), `openxmlpackage-test-${Date.now()}.docx`);
        fs.copyFileSync(srcFile, tmpFile);
        try {
            const buffer = fs.readFileSync(tmpFile);
            const blob = new Blob([buffer]);
            const spy = vi.spyOn(JSZip, 'loadAsync');
            await expect(OpenXmlPackage.open(blob)).resolves.toBeDefined();
            expect(spy).toHaveBeenCalledWith(expect.any(ArrayBuffer));
        } finally {
            fs.unlinkSync(tmpFile);
        }
    });

    it('opens a base64-encoded docx via openFromBase64Internal', async () => {
        const spy = vi.spyOn(JSZip, 'loadAsync');
        await expect(OpenXmlPackage.open(blankDocumentBase64)).resolves.toBeDefined();
        expect(spy).toHaveBeenCalledWith(expect.any(String), { base64: true });
    });

    it('opens a FlatOPC string via openFlatOpcFromXDoc', async () => {
        const spy = vi.spyOn(JSZip, 'loadAsync');
        await expect(OpenXmlPackage.open(blankDocumentFlatOpc)).resolves.toBeDefined();
        expect(spy).not.toHaveBeenCalled();
    });

    it('opens a docx blob with the correct parts', async () => {
        const expectedParts = [
            { uri: '/_rels/.rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
            { uri: '/word/_rels/document.xml.rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
            { uri: '/word/document.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' },
            { uri: '/word/theme/theme1.xml', contentType: 'application/vnd.openxmlformats-officedocument.theme+xml' },
            { uri: '/word/settings.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml' },
            { uri: '/word/fontTable.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml' },
            { uri: '/word/webSettings.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml' },
            { uri: '/docProps/app.xml', contentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' },
            { uri: '/docProps/core.xml', contentType: 'application/vnd.openxmlformats-package.core-properties+xml' },
            { uri: '/word/styles.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' },
        ];
        const srcFile = path.resolve(__dirname, '../../test-files/TemplateDocument.docx');
        const buffer = fs.readFileSync(srcFile);
        const blob = new Blob([buffer]);
        const pkg = await OpenXmlPackage.open(blob);
        const actualParts = pkg.getParts().map(p => ({ uri: p.getUri(), contentType: p.getContentType() }));
        expect(actualParts).toEqual(expectedParts);
    });

    it('opens a base64-encoded docx with the correct parts', async () => {
        const expectedParts = [
            { uri: '/_rels/.rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
            { uri: '/word/_rels/document.xml.rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
            { uri: '/word/document.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' },
            { uri: '/word/theme/theme1.xml', contentType: 'application/vnd.openxmlformats-officedocument.theme+xml' },
            { uri: '/word/settings.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml' },
            { uri: '/word/fontTable.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml' },
            { uri: '/word/webSettings.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml' },
            { uri: '/docProps/app.xml', contentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' },
            { uri: '/docProps/core.xml', contentType: 'application/vnd.openxmlformats-package.core-properties+xml' },
            { uri: '/word/styles.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' },
        ];
        const pkg = await OpenXmlPackage.open(blankDocumentBase64);
        const actualParts = pkg.getParts().map(p => ({ uri: p.getUri(), contentType: p.getContentType() }));
        expect(actualParts).toEqual(expectedParts);
    });

    it('opens a FlatOPC string with the correct parts', async () => {
        const expectedParts = [
            { uri: '/_rels/.rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
            { uri: '/word/document.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' },
            { uri: '/word/_rels/document.xml.rels', contentType: 'application/vnd.openxmlformats-package.relationships+xml' },
            { uri: '/word/theme/theme1.xml', contentType: 'application/vnd.openxmlformats-officedocument.theme+xml' },
            { uri: '/word/settings.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml' },
            { uri: '/word/styles.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' },
            { uri: '/word/webSettings.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml' },
            { uri: '/word/fontTable.xml', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml' },
            { uri: '/docProps/core.xml', contentType: 'application/vnd.openxmlformats-package.core-properties+xml' },
            { uri: '/docProps/app.xml', contentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' },
            { uri: '[Content_Types].xml', contentType: null },
        ];
        const pkg = await OpenXmlPackage.open(blankDocumentFlatOpc);
        const actualParts = pkg.getParts().map(p => ({ uri: p.getUri(), contentType: p.getContentType() }));
        expect(actualParts).toEqual(expectedParts);
    });
});
