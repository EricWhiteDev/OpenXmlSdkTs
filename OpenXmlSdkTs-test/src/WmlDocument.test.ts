import { describe, it, expect } from 'vitest';
import { WmlDocument } from 'OpenXmlSdkTs';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

describe('WmlDocument', () => {
  it('throws when file does not exist', () => {
    expect(() => WmlDocument.open('/nonexistent/path/file.docx')).toThrow();
  });

  it('does not throw when file exists', () => {
    const tmpFile = path.join(os.tmpdir(), `wml-test-${Date.now()}.docx`);
    fs.writeFileSync(tmpFile, '');
    try {
      expect(() => WmlDocument.open(tmpFile)).not.toThrow();
    } finally {
      fs.unlinkSync(tmpFile);
    }
  });
});
