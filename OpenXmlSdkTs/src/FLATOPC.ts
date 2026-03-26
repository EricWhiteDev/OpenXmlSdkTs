import { XName, XNamespace } from 'ltxmlts';

export class FLATOPC {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/2006/xmlPackage");

    static readonly binaryData: XName = FLATOPC.namespace.getName("binaryData");
    static readonly compression: XName = FLATOPC.namespace.getName("compression");
    static readonly contentType: XName = FLATOPC.namespace.getName("contentType");
    static readonly _name: XName = FLATOPC.namespace.getName("name");
    static readonly padding: XName = FLATOPC.namespace.getName("padding");
    static readonly _package: XName = FLATOPC.namespace.getName("package");
    static readonly part: XName = FLATOPC.namespace.getName("part");
    static readonly xmlData: XName = FLATOPC.namespace.getName("xmlData");
}
