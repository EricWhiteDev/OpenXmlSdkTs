import { XName, XNamespace } from 'ltxmlts';

export class CP {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");

    static readonly category: XName = CP.namespace.getName("category");
    static readonly contentStatus: XName = CP.namespace.getName("contentStatus");
    static readonly contentType: XName = CP.namespace.getName("contentType");
    static readonly coreProperties: XName = CP.namespace.getName("coreProperties");
    static readonly keywords: XName = CP.namespace.getName("keywords");
    static readonly lastModifiedBy: XName = CP.namespace.getName("lastModifiedBy");
    static readonly lastPrinted: XName = CP.namespace.getName("lastPrinted");
    static readonly revision: XName = CP.namespace.getName("revision");
}
