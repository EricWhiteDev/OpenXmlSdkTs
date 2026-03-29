/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import {
  XDocument,
  XDeclaration,
  XElement,
  XAttribute,
  XNamespace,
  XProcessingInstruction,
} from "ltxmlts";
import JSZip from "jszip";
import { OpenXmlPart, PartType } from "./OpenXmlPart";
import { OpenXmlRelationship } from "./OpenXmlRelationship";
import { OpenXmlUtility } from "./OpenXmlUtility";
import { CT, FLATOPC, PKGREL } from "./OpenXmlNamespacesAndNames";
import { ContentType } from "./ContentType";
import { RelationshipType } from "./RelationshipType";

export type Base64String = string;
export type FlatOpcString = string;
export type DocxBinary = Blob;

export class OpenXmlPackage {
  private parts: Map<string, OpenXmlPart> = new Map();
  private ctXDoc!: XDocument; // This is the XDocument for the content types in the package

  getParts(): OpenXmlPart[] {
    return Array.from(this.parts.values()).filter(
      (p) =>
        p.getUri() !== "[Content_Types].xml" && p.getContentType() !== ContentType.relationships,
    );
  }

  addPart(uri: string, contentType: string, partType: PartType, data: unknown): OpenXmlPart {
    const alreadyInCt = this.ctXDoc
      .root!.elements(CT.Override)
      .find((el) => el.attribute("PartName")?.value === uri);
    if (alreadyInCt || this.parts.has(uri)) {
      throw new Error(`Invalid operation: part already exists: ${uri}`);
    }
    const newPart = new OpenXmlPart(this, uri, contentType, partType, data);
    this.parts.set(uri, newPart);
    this.ctXDoc.root!.add(
      new XElement(
        CT.Override,
        new XAttribute("PartName", uri),
        new XAttribute("ContentType", contentType),
      ),
    );
    return newPart;
  }

  async deletePart(part: OpenXmlPart): Promise<void> {
    const uri = part.getUri();
    this.parts.delete(uri);

    this.ctXDoc
      .root!.elements(CT.Override)
      .find((el) => el.attribute("PartName")?.value === uri)
      ?.remove();

    const deletedFilename = uri.substring(uri.lastIndexOf("/") + 1);
    const relsDir = uri.substring(0, uri.lastIndexOf("/") + 1) + "_rels/";

    for (const [relsUri, relsPart] of this.parts) {
      if (!relsUri.startsWith(relsDir)) {
        continue;
      }
      let relsXDoc: XDocument;
      const relsData = relsPart.getData();
      if (relsData instanceof XDocument) {
        relsXDoc = relsData;
      } else {
        const xmlStr = await (relsData as { async(type: string): Promise<string> }).async("string");
        relsXDoc = XDocument.parse(xmlStr);
        relsPart.setData(relsXDoc);
        relsPart.setPartType("xml");
      }
      relsXDoc
        .root!.elements()
        .find((el) => el.attribute("Target")?.value === deletedFilename)
        ?.remove();
    }
  }

  getPartByUri(uri: string): OpenXmlPart | undefined {
    return this.parts.get(uri);
  }

  getContentType(uri: string): string {
    return OpenXmlPackage.getContentType(uri, this.ctXDoc);
  }

  async getRelationships(): Promise<OpenXmlRelationship[]> {
    const relsPart = this.parts.get("/_rels/.rels");
    if (!relsPart) {
      return [];
    }
    return OpenXmlPackage.getRelationshipsFromRelsXml(this, null, relsPart);
  }

  async getRelationshipsByRelationshipType(
    relationshipType: string,
  ): Promise<OpenXmlRelationship[]> {
    const rels = await this.getRelationships();
    return rels.filter((r) => r.getType() === relationshipType);
  }

  async getPartsByRelationshipType(relationshipType: string): Promise<OpenXmlPart[]> {
    const rels = await this.getRelationshipsByRelationshipType(relationshipType);
    return rels
      .map((r) => this.getPartByUri(r.getTargetFullName()))
      .filter((p): p is OpenXmlPart => p !== undefined);
  }

  async getRelationshipById(rId: string): Promise<OpenXmlRelationship | undefined> {
    const rels = await this.getRelationships();
    return rels.find((r) => r.getId() === rId);
  }

  async getPartById(rId: string): Promise<OpenXmlPart | undefined> {
    const rel = await this.getRelationshipById(rId);
    return rel ? this.getPartByUri(rel.getTargetFullName()) : undefined;
  }

  async getRelationshipsByContentType(contentType: string): Promise<OpenXmlRelationship[]> {
    const rels = await this.getRelationships();
    return rels.filter(
      (r) =>
        r.getTargetMode() !== "External" &&
        this.getContentType(r.getTargetFullName()) === contentType,
    );
  }

  async getPartsByContentType(contentType: string): Promise<OpenXmlPart[]> {
    const rels = await this.getRelationshipsByContentType(contentType);
    return rels
      .map((r) => this.getPartByUri(r.getTargetFullName()))
      .filter((p): p is OpenXmlPart => p !== undefined);
  }

  async getPartByRelationshipType(relationshipType: string): Promise<OpenXmlPart | undefined> {
    const parts = await this.getPartsByRelationshipType(relationshipType);
    return parts[0];
  }

  async coreFilePropertiesPart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.coreFileProperties);
  }

  async extendedFilePropertiesPart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.extendedFileProperties);
  }

  async customFilePropertiesPart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.customFileProperties);
  }

  async getRelationshipsForPart(part: OpenXmlPart): Promise<OpenXmlRelationship[]> {
    const relsPart = OpenXmlUtility.getRelsPart(part);
    if (!relsPart) {
      return [];
    }
    return OpenXmlPackage.getRelationshipsFromRelsXml(this, part, relsPart);
  }

  async addRelationship(
    id: string,
    type: string,
    target: string,
    targetMode: string = "Internal",
  ): Promise<OpenXmlRelationship> {
    const relsPart = this.getOrCreateRelsPartForUri("/_rels/.rels");
    return OpenXmlPackage.addRelationshipToRelPart(
      this,
      null,
      relsPart,
      id,
      type,
      target,
      targetMode,
    );
  }

  async addRelationshipForPart(
    part: OpenXmlPart,
    id: string,
    type: string,
    target: string,
    targetMode: string = "Internal",
  ): Promise<OpenXmlRelationship> {
    const relsUri = OpenXmlUtility.getRelsPartUri(part);
    const relsPart = this.getOrCreateRelsPartForUri(relsUri);
    return OpenXmlPackage.addRelationshipToRelPart(
      this,
      part,
      relsPart,
      id,
      type,
      target,
      targetMode,
    );
  }

  async deleteRelationship(id: string): Promise<boolean> {
    const relsPart = this.parts.get("/_rels/.rels");
    if (!relsPart) {
      throw new Error(`Relationship not found: ${id}`);
    }
    return OpenXmlPackage.deleteRelationshipFromRelPart(relsPart, id);
  }

  async deleteRelationshipForPart(part: OpenXmlPart, id: string): Promise<boolean> {
    const relsPart = OpenXmlUtility.getRelsPart(part);
    if (!relsPart) {
      throw new Error(`Relationship not found: ${id}`);
    }
    return OpenXmlPackage.deleteRelationshipFromRelPart(relsPart, id);
  }

  async saveToFlatOpcAsync(): Promise<FlatOpcString> {
    const pkgElement = new XElement(
      FLATOPC._package,
      new XAttribute(XNamespace.xmlns.getName("pkg"), FLATOPC.namespace.namespaceName),
    );
    const flatOpc = new XDocument(
      new XDeclaration("1.0", "UTF-8", "yes"),
      new XProcessingInstruction("mso-application", 'progid="Word.Document"'),
      pkgElement,
    );

    for (const [uri, part] of this.parts) {
      if (uri === "[Content_Types].xml") {
        continue;
      }

      const ct = part.getContentType()!;
      const partType = part.getPartType();
      const data = part.getData();

      let compression: string | null = null;
      let dataElement: XElement;

      if (partType === "xml") {
        let root: unknown;
        if (data instanceof XDocument) {
          root = data.root;
        } else {
          const xmlStr = await (data as { async(type: string): Promise<string> }).async("string");
          root = XDocument.parse(xmlStr).root;
        }
        dataElement = new XElement(FLATOPC.xmlData, root);
      } else {
        let content: string;
        if (typeof data === "string") {
          content = data;
        } else {
          content = await (data as { async(type: string): Promise<string> }).async("base64");
        }
        compression = "store";
        dataElement = new XElement(FLATOPC.binaryData, content);
      }

      const partElement = new XElement(
        FLATOPC.part,
        new XAttribute(FLATOPC._name, uri),
        new XAttribute(FLATOPC.contentType, ct),
        compression ? new XAttribute(FLATOPC.compression, compression) : null,
        dataElement,
      );
      pkgElement.add(partElement);
    }

    return flatOpc.toStringWithIndentation() as FlatOpcString;
  }

  async saveToBase64Async(): Promise<Base64String> {
    const zip = await this.saveToZip();
    return zip.generateAsync({ type: "base64", compression: "DEFLATE" });
  }

  async saveToBlobAsync(): Promise<DocxBinary> {
    const zip = await this.saveToZip();
    return zip.generateAsync({ type: "blob", compression: "DEFLATE" });
  }

  private getOrCreateRelsPartForUri(relsUri: string): OpenXmlPart {
    const existing = this.parts.get(relsUri);
    if (existing) {
      return existing;
    }
    const relsXDoc = new XDocument(
      new XElement(PKGREL.Relationships, new XAttribute("xmlns", PKGREL.namespace.namespaceName)),
    );
    const newPart = new OpenXmlPart(this, relsUri, ContentType.relationships, "xml", relsXDoc);
    this.parts.set(relsUri, newPart);
    return newPart;
  }

  private static async deleteRelationshipFromRelPart(
    relsPart: OpenXmlPart,
    id: string,
  ): Promise<boolean> {
    let relsXDoc: XDocument;
    const data = relsPart.getData();
    if (data instanceof XDocument) {
      relsXDoc = data;
    } else {
      const xmlStr = await (data as { async(type: string): Promise<string> }).async("string");
      relsXDoc = XDocument.parse(xmlStr);
      relsPart.setData(relsXDoc);
      relsPart.setPartType("xml");
    }
    const el = relsXDoc
      .root!.elements(PKGREL.Relationship)
      .find((r) => r.attribute("Id")?.value === id);
    if (!el) {
      throw new Error(`Relationship not found: ${id}`);
    }
    el.remove();
    return true;
  }

  private static async addRelationshipToRelPart(
    pkg: OpenXmlPackage,
    part: OpenXmlPart | null,
    relsPart: OpenXmlPart,
    id: string,
    type: string,
    target: string,
    targetMode: string,
  ): Promise<OpenXmlRelationship> {
    let relsXDoc: XDocument;
    const data = relsPart.getData();
    if (data instanceof XDocument) {
      relsXDoc = data;
    } else {
      const xmlStr = await (data as { async(type: string): Promise<string> }).async("string");
      relsXDoc = XDocument.parse(xmlStr);
      relsPart.setData(relsXDoc);
      relsPart.setPartType("xml");
    }
    relsXDoc.root!.add(
      new XElement(
        PKGREL.Relationship,
        new XAttribute("Id", id),
        new XAttribute("Type", type),
        new XAttribute("Target", target),
        targetMode !== "Internal" ? new XAttribute("TargetMode", targetMode) : null,
      ),
    );
    const storedTargetMode = targetMode !== "Internal" ? targetMode : null;
    return new OpenXmlRelationship(pkg, part, id, type, target, storedTargetMode);
  }

  private static async getRelationshipsFromRelsXml(
    pkg: OpenXmlPackage,
    part: OpenXmlPart | null,
    relsPart: OpenXmlPart,
  ): Promise<OpenXmlRelationship[]> {
    let relsXDoc: XDocument;
    const data = relsPart.getData();
    if (data instanceof XDocument) {
      relsXDoc = data;
    } else {
      const xmlStr = await (data as { async(type: string): Promise<string> }).async("string");
      relsXDoc = XDocument.parse(xmlStr);
    }
    return relsXDoc.root!.elements(PKGREL.Relationship).map((r) => {
      const targetMode = r.attribute("TargetMode")?.value ?? null;
      return new OpenXmlRelationship(
        pkg,
        part,
        r.attribute("Id")!.value,
        r.attribute("Type")!.value,
        r.attribute("Target")!.value,
        targetMode,
      );
    });
  }

  private async saveToZip(): Promise<JSZip> {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", this.ctXDoc.toString());

    for (const [uri, part] of this.parts) {
      if (uri === "[Content_Types].xml") {
        continue;
      }

      const name = uri.startsWith("/") ? uri.substring(1) : uri;
      const partType = part.getPartType();
      const data = part.getData();

      if (partType === "xml") {
        if (data instanceof XDocument) {
          zip.file(name, data.toString());
        } else {
          const xmlStr = await (data as { async(type: string): Promise<string> }).async("string");
          zip.file(name, xmlStr);
        }
      } else {
        if (typeof data === "string") {
          zip.file(name, data, { base64: true });
        } else {
          const bytes = await (data as { async(type: string): Promise<Uint8Array> }).async(
            "uint8array",
          );
          zip.file(name, bytes);
        }
      }
    }

    return zip;
  }

  protected static async openInto<T extends OpenXmlPackage>(
    pkg: T,
    document: Base64String | FlatOpcString | DocxBinary,
  ): Promise<T> {
    if (typeof document === "string") {
      if (OpenXmlUtility.isBase64(document)) {
        await OpenXmlPackage.openFromBase64Internal(pkg, document);
      } else {
        await OpenXmlPackage.openFromFlatOpcInternal(pkg, document);
      }
    } else if (document instanceof Blob) {
      await OpenXmlPackage.openFromBlobInternal(pkg, document);
    } else {
      throw new Error(
        "Invalid argument: document must be a Base64String, FlatOpcString, or DocxBinary (Blob).",
      );
    }
    return pkg;
  }

  static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<OpenXmlPackage> {
    return OpenXmlPackage.openInto(new OpenXmlPackage(), document);
  }

  private static getContentType(uri: string, ctXDoc: XDocument): string {
    const root = ctXDoc.root!;

    const override = root
      .elements(CT.Override)
      .find((el) => el.attribute("PartName")?.value === uri);
    if (override) {
      const ct = override.attribute("ContentType")?.value;
      if (ct) {
        return ct;
      }
    }

    const ext = uri.split(".").pop() ?? "";
    const def = root.elements(CT.Default).find((el) => el.attribute("Extension")?.value === ext);
    if (def) {
      const ct = def.attribute("ContentType")?.value;
      if (ct) {
        return ct;
      }
    }

    throw new Error(`Content type not found for part: ${uri}`);
  }

  private static async openFromBase64Internal(
    pkg: OpenXmlPackage,
    document: Base64String,
  ): Promise<void> {
    const zip = await JSZip.loadAsync(document, { base64: true });
    await OpenXmlPackage.openFromZip(zip, pkg);
  }

  private static openFlatOpcFromXDoc(pkg: OpenXmlPackage, doc: XDocument): void {
    const root = doc.root!;
    pkg.ctXDoc = new XDocument(
      new XDeclaration("1.0", "utf-8", "yes"),
      new XElement(
        CT.Types,
        new XAttribute("xmlns", CT.namespace.namespaceName),
        new XElement(
          CT.Default,
          new XAttribute("Extension", "rels"),
          new XAttribute("ContentType", ContentType.relationships),
        ),
        new XElement(
          CT.Default,
          new XAttribute("Extension", "xml"),
          new XAttribute("ContentType", "application/xml"),
        ),
      ),
    );

    for (const p of root.elements(FLATOPC.part)) {
      const uri = p.attribute(FLATOPC._name)!.value;
      const contentType = p.attribute(FLATOPC.contentType)!.value;
      const partType = contentType.endsWith("xml") ? "xml" : "base64";

      if (partType === "xml") {
        const xmlDataEl = p.element(FLATOPC.xmlData)!;
        const newPart = new OpenXmlPart(
          pkg,
          uri,
          contentType,
          "xml",
          new XDocument(xmlDataEl.elements()[0]),
        );
        pkg.parts.set(uri, newPart);
        if (contentType !== ContentType.relationships) {
          pkg.ctXDoc.root!.add(
            new XElement(
              CT.Override,
              new XAttribute("PartName", uri),
              new XAttribute("ContentType", contentType),
            ),
          );
        }
      } else {
        const binaryData = p.element(FLATOPC.binaryData)!.value;
        const newPart = new OpenXmlPart(pkg, uri, contentType, "binary", binaryData);
        pkg.parts.set(uri, newPart);
        pkg.ctXDoc.root!.add(
          new XElement(
            CT.Override,
            new XAttribute("PartName", uri),
            new XAttribute("ContentType", contentType),
          ),
        );
      }
    }

    const ctPart = new OpenXmlPart(pkg, "[Content_Types].xml", null, "xml", pkg.ctXDoc);
    pkg.parts.set("[Content_Types].xml", ctPart);
  }

  private static async openFromFlatOpcInternal(
    pkg: OpenXmlPackage,
    document: FlatOpcString,
  ): Promise<void> {
    const xDoc = XDocument.parse(document);
    OpenXmlPackage.openFlatOpcFromXDoc(pkg, xDoc);
  }

  private static async openFromBlobInternal(
    pkg: OpenXmlPackage,
    document: DocxBinary,
  ): Promise<void> {
    const arrayBuffer = await document.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    await OpenXmlPackage.openFromZip(zip, pkg);
  }

  private static async openFromZip(zip: JSZip, pkg: OpenXmlPackage): Promise<void> {
    const ctZipFile = zip.files["[Content_Types].xml"];
    if (!ctZipFile) {
      throw new Error("Invalid Open XML document: no [Content_Types].xml");
    }
    const ctData = await ctZipFile.async("string");
    pkg.ctXDoc = XDocument.parse(ctData);

    for (const f in zip.files) {
      const zipFile = zip.files[f];
      if (!f.endsWith("/") && f !== "[Content_Types].xml") {
        const f2 = "/" + f;
        const newPart = new OpenXmlPart(pkg, f2, null, null, zipFile);
        pkg.parts.set(f2, newPart);
      }
    }

    for (const [part, thisPart] of pkg.parts) {
      const ct = OpenXmlPackage.getContentType(part, pkg.ctXDoc);
      thisPart.setContentType(ct);
      thisPart.setPartType(ct.endsWith("xml") ? "xml" : "binary");
    }
  }
}
