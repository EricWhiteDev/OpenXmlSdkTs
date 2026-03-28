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
});
