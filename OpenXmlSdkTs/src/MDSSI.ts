import { XName, XNamespace } from 'ltxmlts';

export class MDSSI {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/package/2006/digital-signature");

    static readonly Format: XName = MDSSI.namespace.getName("Format");
    static readonly RelationshipReference: XName = MDSSI.namespace.getName("RelationshipReference");
    static readonly SignatureTime: XName = MDSSI.namespace.getName("SignatureTime");
    static readonly Value: XName = MDSSI.namespace.getName("Value");
}
