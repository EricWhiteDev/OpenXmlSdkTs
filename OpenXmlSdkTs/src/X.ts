import { XName, XNamespace } from 'ltxmlts';

export class X {
    static readonly namespace: XNamespace = XNamespace.get("urn:schemas-microsoft-com:office:excel");

    static readonly Anchor: XName = X.namespace.getName("Anchor");
    static readonly AutoFill: XName = X.namespace.getName("AutoFill");
    static readonly ClientData: XName = X.namespace.getName("ClientData");
    static readonly Column: XName = X.namespace.getName("Column");
    static readonly MoveWithCells: XName = X.namespace.getName("MoveWithCells");
    static readonly Row: XName = X.namespace.getName("Row");
    static readonly SizeWithCells: XName = X.namespace.getName("SizeWithCells");
}
