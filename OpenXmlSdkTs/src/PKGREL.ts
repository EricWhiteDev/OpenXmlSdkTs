import { XName, XNamespace } from 'ltxmlts';

export class PKGREL {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/package/2006/relationships");

    static readonly Relationship: XName = PKGREL.namespace.getName("Relationship");
    static readonly Relationships: XName = PKGREL.namespace.getName("Relationships");
}
