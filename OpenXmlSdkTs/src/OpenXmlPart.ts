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
import { OpenXmlUtility } from "./OpenXmlUtility";
import { RelationshipType } from "./RelationshipType";

export type PartType = "binary" | "base64" | "xml" | null;

export class OpenXmlPart {
  private pkg: OpenXmlPackage;
  private uri: string;
  private contentType: string | null; // MIME type value, e.g. "application/vnd...+xml"
  private partType: PartType;
  private data: unknown; // for now, this type is unknown.  May change later.

  constructor(
    pkg: OpenXmlPackage,
    uri: string,
    contentType: string | null,
    partType: PartType,
    data: unknown,
  ) {
    this.pkg = pkg;
    this.uri = uri;
    this.contentType = contentType;
    this.partType = partType;
    this.data = data;
  }

  getUri(): string {
    return this.uri;
  }

  getPkg(): OpenXmlPackage {
    return this.pkg;
  }

  getContentType(): string | null {
    return this.contentType;
  }

  getData(): unknown {
    return this.data;
  }

  setContentType(ct: string): void {
    this.contentType = ct;
  }

  getPartType(): PartType {
    return this.partType;
  }

  setData(data: unknown): void {
    this.data = data;
  }

  setPartType(pt: PartType): void {
    this.partType = pt;
  }

  getRelsPartUri(): string {
    return OpenXmlUtility.getRelsPartUri(this);
  }

  getRelsPart(): OpenXmlPart | undefined {
    return OpenXmlUtility.getRelsPart(this);
  }

  async getRelationships(): Promise<OpenXmlRelationship[]> {
    return this.pkg.getRelationshipsForPart(this);
  }

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

  async getRelationshipsByRelationshipType(
    relationshipType: string,
  ): Promise<OpenXmlRelationship[]> {
    const rels = await this.getRelationships();
    return rels.filter((r) => r.getType() === relationshipType);
  }

  async getPartsByRelationshipType(relationshipType: string): Promise<OpenXmlPart[]> {
    const rels = await this.getRelationshipsByRelationshipType(relationshipType);
    return rels
      .map((r) => this.pkg.getPartByUri(r.getTargetFullName()))
      .filter((p): p is OpenXmlPart => p !== undefined);
  }

  async getPartByRelationshipType(relationshipType: string): Promise<OpenXmlPart | undefined> {
    const parts = await this.getPartsByRelationshipType(relationshipType);
    return parts[0];
  }

  async addRelationship(
    id: string,
    type: string,
    target: string,
    targetMode: string = "Internal",
  ): Promise<OpenXmlRelationship> {
    return this.pkg.addRelationshipForPart(this, id, type, target, targetMode);
  }

  async deleteRelationship(id: string): Promise<boolean> {
    return this.pkg.deleteRelationshipForPart(this, id);
  }

  async getXDocument(): Promise<XDocument> {
    if (this.partType !== "xml") {
      throw new Error(`Cannot get XDocument for non-xml part: ${this.uri}`);
    }
    if (this.data instanceof XDocument) {
      return this.data;
    }
    const xmlStr = await (this.data as { async(type: string): Promise<string> }).async("string");
    const xDoc = XDocument.parse(xmlStr);
    this.data = xDoc;
    return xDoc;
  }

  putXDocument(xDoc: XDocument): void {
    if (!xDoc) {
      throw new Error("putXDocument: xDoc must not be null or undefined");
    }
    this.data = xDoc;
    this.partType = "xml";
  }

  async headerParts(): Promise<OpenXmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.header);
  }

  async footerParts(): Promise<OpenXmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.footer);
  }

  async endnotesPart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.endnotes);
  }

  async footnotesPart(): Promise<OpenXmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.footnotes);
  }

  async getRelationshipById(rId: string): Promise<OpenXmlRelationship | undefined> {
    const rels = await this.getRelationships();
    return rels.find((r) => r.getId() === rId);
  }

  async getPartById(rId: string): Promise<OpenXmlPart | undefined> {
    const rel = await this.getRelationshipById(rId);
    return rel ? this.pkg.getPartByUri(rel.getTargetFullName()) : undefined;
  }

  async getRelationshipsByContentType(contentType: string): Promise<OpenXmlRelationship[]> {
    const rels = await this.getRelationships();
    return rels.filter(
      (r) =>
        r.getTargetMode() !== "External" &&
        this.pkg.getContentType(r.getTargetFullName()) === contentType,
    );
  }

  async getPartsByContentType(contentType: string): Promise<OpenXmlPart[]> {
    const rels = await this.getRelationshipsByContentType(contentType);
    return rels
      .map((r) => this.pkg.getPartByUri(r.getTargetFullName()))
      .filter((p): p is OpenXmlPart => p !== undefined);
  }
}
