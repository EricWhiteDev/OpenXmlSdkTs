import { XName, XNamespace } from 'ltxmlts';

export class VT {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

    static readonly _bool: XName = VT.namespace.getName("bool");
    static readonly filetime: XName = VT.namespace.getName("filetime");
    static readonly i4: XName = VT.namespace.getName("i4");
    static readonly lpstr: XName = VT.namespace.getName("lpstr");
    static readonly lpwstr: XName = VT.namespace.getName("lpwstr");
    static readonly r8: XName = VT.namespace.getName("r8");
    static readonly variant: XName = VT.namespace.getName("variant");
    static readonly vector: XName = VT.namespace.getName("vector");
}
