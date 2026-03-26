import { XName, XNamespace } from 'ltxmlts';

export class CT {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/package/2006/content-types");

    static readonly Default: XName = CT.namespace.getName("Default");
    static readonly Override: XName = CT.namespace.getName("Override");
    static readonly Types: XName = CT.namespace.getName("Types");
}
