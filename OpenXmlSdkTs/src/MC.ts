import { XName, XNamespace } from 'ltxmlts';

export class MC {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/markup-compatibility/2006");

    static readonly AlternateContent: XName = MC.namespace.getName("AlternateContent");
    static readonly Choice: XName = MC.namespace.getName("Choice");
    static readonly Fallback: XName = MC.namespace.getName("Fallback");
    static readonly Ignorable: XName = MC.namespace.getName("Ignorable");
    static readonly PreserveAttributes: XName = MC.namespace.getName("PreserveAttributes");
}
