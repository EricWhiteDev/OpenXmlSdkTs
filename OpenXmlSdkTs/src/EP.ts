import { XName, XNamespace } from 'ltxmlts';

export class EP {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");

    static readonly Application: XName = EP.namespace.getName("Application");
    static readonly AppVersion: XName = EP.namespace.getName("AppVersion");
    static readonly Characters: XName = EP.namespace.getName("Characters");
    static readonly CharactersWithSpaces: XName = EP.namespace.getName("CharactersWithSpaces");
    static readonly Company: XName = EP.namespace.getName("Company");
    static readonly DocSecurity: XName = EP.namespace.getName("DocSecurity");
    static readonly HeadingPairs: XName = EP.namespace.getName("HeadingPairs");
    static readonly HiddenSlides: XName = EP.namespace.getName("HiddenSlides");
    static readonly HLinks: XName = EP.namespace.getName("HLinks");
    static readonly HyperlinkBase: XName = EP.namespace.getName("HyperlinkBase");
    static readonly HyperlinksChanged: XName = EP.namespace.getName("HyperlinksChanged");
    static readonly Lines: XName = EP.namespace.getName("Lines");
    static readonly LinksUpToDate: XName = EP.namespace.getName("LinksUpToDate");
    static readonly Manager: XName = EP.namespace.getName("Manager");
    static readonly MMClips: XName = EP.namespace.getName("MMClips");
    static readonly Notes: XName = EP.namespace.getName("Notes");
    static readonly Pages: XName = EP.namespace.getName("Pages");
    static readonly Paragraphs: XName = EP.namespace.getName("Paragraphs");
    static readonly PresentationFormat: XName = EP.namespace.getName("PresentationFormat");
    static readonly Properties: XName = EP.namespace.getName("Properties");
    static readonly ScaleCrop: XName = EP.namespace.getName("ScaleCrop");
    static readonly SharedDoc: XName = EP.namespace.getName("SharedDoc");
    static readonly Slides: XName = EP.namespace.getName("Slides");
    static readonly Template: XName = EP.namespace.getName("Template");
    static readonly TitlesOfParts: XName = EP.namespace.getName("TitlesOfParts");
    static readonly TotalTime: XName = EP.namespace.getName("TotalTime");
    static readonly Words: XName = EP.namespace.getName("Words");
}
