import { XName, XNamespace } from 'ltxmlts';

export class DCTERMS {
    static readonly namespace: XNamespace = XNamespace.get("http://purl.org/dc/terms/");

    static readonly created: XName = DCTERMS.namespace.getName("created");
    static readonly modified: XName = DCTERMS.namespace.getName("modified");
}
