/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { XDocument } from "ltxmlts";
import { OpenXmlPackage } from "./OpenXmlPackage";
import { OpenXmlRelationship } from "./OpenXmlRelationship";
import { Utility } from "./Utility";
import { RelationshipType } from "./RelationshipType";
import { parseXmlPreservingWhitespace } from "./XmlParsing";

/**
 * Describes the data format of a part's content.
 *
 * - `"xml"` — The part contains XML data (accessible via {@link OpenXmlPart.getXDocument}).
 * - `"binary"` — The part contains raw binary data.
 * - `"base64"` — The part contains Base64-encoded binary data.
 * - `null` — The part type has not yet been determined.
 *
 * @category Class and Type Reference
 */
export type PartType = "binary" | "base64" | "xml" | null;

/**
 * Represents a single part within an Open XML package.
 *
 * @remarks
 * A part is one file inside the ZIP archive that makes up an Open XML document.
 * Parts can contain XML (document content, styles, relationships) or binary data
 * (images, embedded objects).
 *
 * Use {@link OpenXmlPart.getXDocument | getXDocument()} to read and
 * {@link OpenXmlPart.putXDocument | putXDocument()} to write XML content.
 * Navigate to related parts via the relationship methods.
 *
 * For format-specific convenience methods, see the subclasses
 * {@link WmlPart}, {@link SmlPart}, and {@link PmlPart}.
 *
 * @example
 * ```typescript
 * const mainPart = await doc.mainDocumentPart();
 * const xDoc = await mainPart!.getXDocument();
 * const body = xDoc.root!.element(W.body);
 * console.log(`${body!.elements(W.p).length} paragraphs`);
 *
 * // Navigate to related parts
 * const stylesPart = await mainPart!.getPartByRelationshipType(RelationshipType.styles);
 * ```
 *
 * @category Class and Type Reference
 */
export class OpenXmlPart {
  private pkg: OpenXmlPackage;
  private uri: string;
  private contentType: string | null;
  private partType: PartType;
  private data: unknown;

  /** @internal */
  constructor(pkg: OpenXmlPackage, uri: string, contentType: string | null, partType: PartType, data: unknown) {
    this.pkg = pkg;
    this.uri = uri;
    this.contentType = contentType;
    this.partType = partType;
    this.data = data;
  }

  /** Returns the URI of this part within the package (e.g., `"/word/document.xml"`). */
  getUri(): string {
    return this.uri;
  }

  /** Returns the {@link OpenXmlPackage} that contains this part. */
  getPkg(): OpenXmlPackage {
    return this.pkg;
  }

  /** Returns the MIME content type of this part (e.g., `"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"`). */
  getContentType(): string | null {
    return this.contentType;
  }

  /** Returns the raw data of this part. For XML parts, prefer {@link getXDocument}. */
  getData(): unknown {
    return this.data;
  }

  /** Sets the MIME content type of this part. */
  setContentType(ct: string): void {
    this.contentType = ct;
  }

  /** Returns the data format of this part. See {@link PartType}. */
  getPartType(): PartType {
    return this.partType;
  }

  /** Sets the raw data of this part. For XML parts, prefer {@link putXDocument}. */
  setData(data: unknown): void {
    this.data = data;
  }

  /** Sets the data format of this part. See {@link PartType}. */
  setPartType(pt: PartType): void {
    this.partType = pt;
  }

  /** Returns the URI of the `.rels` file associated with this part. */
  getRelsPartUri(): string {
    return Utility.getRelsPartUri(this);
  }

  /** Returns the `.rels` part associated with this part, or `undefined` if none exists. */
  getRelsPart(): OpenXmlPart | undefined {
    return Utility.getRelsPart(this);
  }

  /**
   * Returns all relationships defined for this part.
   *
   * @returns An array of {@link OpenXmlRelationship} objects.
   */
  async getRelationships(): Promise<OpenXmlRelationship[]> {
    return this.pkg.getRelationshipsForPart(this);
  }

  /**
   * Returns all internal parts referenced by this part's relationships.
   *
   * @returns An array of {@link OpenXmlPart} instances.
   */
  async getParts(): Promise<OpenXmlPart[]> {
    const rels = await this.getRelationships();
    return rels
      .filter((r) => r.getTargetMode() !== "External")
      .map((r) => {
        const part = this.pkg.getPartByUri(r.getTargetFullName());
        if (!part) {
          throw new Error(`Part not found for relationship target: ${r.getTargetFullName()}`);
        }
        return part;
      });
  }

  /**
   * Returns relationships of this part filtered by relationship type.
   *
   * @param relationshipType - The relationship type URI. Use {@link RelationshipType} constants.
   */
  async getRelationshipsByRelationshipType(relationshipType: string): Promise<OpenXmlRelationship[]> {
    const rels = await this.getRelationships();
    return rels.filter((r) => r.getType() === relationshipType);
  }

  /**
   * Returns parts targeted by relationships of the given type.
   *
   * @param relationshipType - The relationship type URI. Use {@link RelationshipType} constants.
   */
  async getPartsByRelationshipType(relationshipType: string): Promise<OpenXmlPart[]> {
    const rels = await this.getRelationshipsByRelationshipType(relationshipType);
    return rels.map((r) => this.pkg.getPartByUri(r.getTargetFullName())).filter((p): p is OpenXmlPart => p !== undefined);
  }

  /**
   * Returns the first part targeted by a relationship of the given type.
   *
   * @param relationshipType - The relationship type URI. Use {@link RelationshipType} constants.
   * @returns The first matching {@link OpenXmlPart}, or `undefined`.
   */
  async getPartByRelationshipType(relationshipType: string): Promise<OpenXmlPart | undefined> {
    const parts = await this.getPartsByRelationshipType(relationshipType);
    return parts[0];
  }

  /**
   * Adds a relationship to this part.
   *
   * @param id - The relationship ID (e.g., `"rId10"`).
   * @param type - The relationship type URI. Use {@link RelationshipType} constants.
   * @param target - The target URI (relative to this part).
   * @param targetMode - `"Internal"` (default) or `"External"`.
   * @returns The newly created {@link OpenXmlRelationship}.
   */
  async addRelationship(id: string, type: string, target: string, targetMode: string = "Internal"): Promise<OpenXmlRelationship> {
    return this.pkg.addRelationshipForPart(this, id, type, target, targetMode);
  }

  /**
   * Deletes a relationship from this part.
   *
   * @param id - The relationship ID to delete.
   * @returns `true` if the relationship was deleted.
   * @throws Error if the relationship is not found.
   */
  async deleteRelationship(id: string): Promise<boolean> {
    return this.pkg.deleteRelationshipForPart(this, id);
  }

  /**
   * Returns the XML content of this part as an XDocument.
   *
   * @remarks
   * The XDocument is parsed on first access and cached for subsequent calls.
   * Modifications to the returned XDocument are reflected in the part.
   * Call {@link putXDocument} to replace the XML content entirely.
   *
   * @returns A promise resolving to the part's {@link XDocument}.
   * @throws Error if the part is not an XML part.
   *
   * @example
   * ```typescript
   * const xDoc = await mainPart.getXDocument();
   * const body = xDoc.root!.element(W.body);
   * const paragraphs = body!.elements(W.p);
   * ```
   */
  async getXDocument(): Promise<XDocument> {
    if (this.partType !== "xml") {
      throw new Error(`Cannot get XDocument for non-xml part: ${this.uri}`);
    }
    if (this.data instanceof XDocument) {
      return this.data;
    }
    const xmlStr = await (this.data as { async(type: string): Promise<string> }).async("string");
    const xDoc = parseXmlPreservingWhitespace(xmlStr);
    this.data = xDoc;
    return xDoc;
  }

  /**
   * Replaces the XML content of this part.
   *
   * @param xDoc - The new {@link XDocument} content.
   * @throws Error if `xDoc` is null or undefined.
   *
   * @example
   * ```typescript
   * const xDoc = await part.getXDocument();
   * // ... modify xDoc ...
   * part.putXDocument(xDoc);
   * ```
   */
  putXDocument(xDoc: XDocument): void {
    if (!xDoc) {
      throw new Error("putXDocument: xDoc must not be null or undefined");
    }
    this.data = xDoc;
    this.partType = "xml";
  }

  /** Returns the custom XML properties part, or `undefined` if not present. */
  async customXmlPropertiesPart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.customXmlProperties);
  }

  /** Returns the theme part, or `undefined` if not present. */
  async themePart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.theme);
  }

  /** Returns the thumbnail part, or `undefined` if not present. */
  async thumbnailPart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.thumbnail);
  }

  /** Returns the drawings part, or `undefined` if not present. */
  async drawingsPart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.drawings);
  }

  /** Returns all image parts referenced by this part. */
  async imageParts(): Promise<OpenXmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.image);
  }

  /** Returns all custom XML parts referenced by this part. */
  async customXmlParts(): Promise<OpenXmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.customXml);
  }

  /**
   * Finds a relationship of this part by its ID.
   *
   * @param rId - The relationship ID (e.g., `"rId1"`).
   */
  async getRelationshipById(rId: string): Promise<OpenXmlRelationship | undefined> {
    const rels = await this.getRelationships();
    return rels.find((r) => r.getId() === rId);
  }

  /**
   * Finds a part by following a relationship ID from this part.
   *
   * @param rId - The relationship ID (e.g., `"rId1"`).
   */
  async getPartById(rId: string): Promise<OpenXmlPart | undefined> {
    const rel = await this.getRelationshipById(rId);
    return rel ? this.pkg.getPartByUri(rel.getTargetFullName()) : undefined;
  }

  /**
   * Returns relationships whose target parts have the given content type.
   *
   * @param contentType - The MIME content type. Use {@link ContentType} constants.
   */
  async getRelationshipsByContentType(contentType: string): Promise<OpenXmlRelationship[]> {
    const rels = await this.getRelationships();
    return rels.filter((r) => r.getTargetMode() !== "External" && this.pkg.getContentType(r.getTargetFullName()) === contentType);
  }

  /**
   * Returns parts whose content type matches the given value.
   *
   * @param contentType - The MIME content type. Use {@link ContentType} constants.
   */
  async getPartsByContentType(contentType: string): Promise<OpenXmlPart[]> {
    const rels = await this.getRelationshipsByContentType(contentType);
    return rels.map((r) => this.pkg.getPartByUri(r.getTargetFullName())).filter((p): p is OpenXmlPart => p !== undefined);
  }
}
