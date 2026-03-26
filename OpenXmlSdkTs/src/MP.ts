import { XName, XNamespace } from 'ltxmlts';

export class MP {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/mac/powerpoint/2008/main");

    static readonly cube: XName = MP.namespace.getName("cube");
    static readonly flip: XName = MP.namespace.getName("flip");
    static readonly transition: XName = MP.namespace.getName("transition");
}
