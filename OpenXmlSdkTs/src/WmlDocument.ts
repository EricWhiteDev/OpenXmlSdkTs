import * as fs from 'fs';

export class WmlDocument {
  private constructor(private readonly filePath: string) {}

  static open(fileName: string): WmlDocument {
    if (!fs.existsSync(fileName)) {
      throw new Error(`File not found: ${fileName}`);
    }
    return new WmlDocument(fileName);
  }
}
