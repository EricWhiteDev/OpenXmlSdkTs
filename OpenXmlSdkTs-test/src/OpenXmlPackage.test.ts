/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { describe, it, expect } from 'vitest';
import { OpenXmlPackage } from 'OpenXmlSdkTs';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

describe('OpenXmlPackage', () => {
    it('does not throw when opening a docx blob', async () => {
        const srcFile = path.resolve(__dirname, '../../test-files/New Microsoft Word Document.docx');
        const tmpFile = path.join(os.tmpdir(), `openxmlpackage-test-${Date.now()}.docx`);
        fs.copyFileSync(srcFile, tmpFile);
        try {
            const buffer = fs.readFileSync(tmpFile);
            const blob = new Blob([buffer]);
            await expect(OpenXmlPackage.open(blob)).resolves.toBeDefined();
        } finally {
            fs.unlinkSync(tmpFile);
        }
    });
});
