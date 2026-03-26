import { XName, XNamespace } from 'ltxmlts';

export class SL {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/schemaLibrary/2006/main");

    static readonly manifestLocation: XName = SL.namespace.getName("manifestLocation");
    static readonly schema: XName = SL.namespace.getName("schema");
    static readonly schemaLibrary: XName = SL.namespace.getName("schemaLibrary");
    static readonly uri: XName = SL.namespace.getName("uri");
}
