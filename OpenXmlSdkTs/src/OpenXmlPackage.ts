import { XDocument } from 'ltxmlts';
import { OpenXmlPart } from './OpenXmlPart';

export class OpenXmlPackage {
  private parts: Map<string, OpenXmlPart> = new Map();
  // This is the XDocument for the content types in the package
  private ctXDoc!: XDocument;
}
