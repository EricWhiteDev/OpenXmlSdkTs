import { XName, XNamespace } from 'ltxmlts';

export class COM {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/drawingml/2006/compatibility");

    static readonly legacyDrawing: XName = COM.namespace.getName("legacyDrawing");
}
