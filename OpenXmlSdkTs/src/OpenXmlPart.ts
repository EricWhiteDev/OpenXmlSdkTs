/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { OpenXmlPackage } from "./OpenXmlPackage";
import { OpenXmlRelationship } from "./OpenXmlRelationship";

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

  async getRelationships(): Promise<OpenXmlRelationship[]> {
    return this.pkg.getRelationshipsForPart(this);
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
        this.pkg.getContentTypeForUri(r.getTargetFullName()) === contentType,
    );
  }

  async getPartsByContentType(contentType: string): Promise<OpenXmlPart[]> {
    const rels = await this.getRelationshipsByContentType(contentType);
    return rels
      .map((r) => this.pkg.getPartByUri(r.getTargetFullName()))
      .filter((p): p is OpenXmlPart => p !== undefined);
  }
}
