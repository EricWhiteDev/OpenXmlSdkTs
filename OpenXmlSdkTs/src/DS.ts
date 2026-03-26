import { XName, XNamespace } from 'ltxmlts';

export class DS {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/officeDocument/2006/customXml");

    static readonly datastoreItem: XName = DS.namespace.getName("datastoreItem");
    static readonly itemID: XName = DS.namespace.getName("itemID");
    static readonly schemaRef: XName = DS.namespace.getName("schemaRef");
    static readonly schemaRefs: XName = DS.namespace.getName("schemaRefs");
    static readonly uri: XName = DS.namespace.getName("uri");
}
