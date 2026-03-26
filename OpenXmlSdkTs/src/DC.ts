import { XName, XNamespace } from 'ltxmlts';

export class DC {
    static readonly namespace: XNamespace = XNamespace.get("http://purl.org/dc/elements/1.1/");

    static readonly creator: XName = DC.namespace.getName("creator");
    static readonly description: XName = DC.namespace.getName("description");
    static readonly subject: XName = DC.namespace.getName("subject");
    static readonly title: XName = DC.namespace.getName("title");
}
