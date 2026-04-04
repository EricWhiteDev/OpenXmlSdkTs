/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import type { OpenXmlPackage } from "./OpenXmlPackage";
import type { OpenXmlPart } from "./OpenXmlPart";

/**
 * Represents a single relationship within an Open XML package.
 *
 * @remarks
 * Relationships link parts to other parts or external resources. Each relationship
 * has a unique ID, a type URI, and a target URI. Package-level relationships are
 * stored in `/_rels/.rels`; part-level relationships are stored alongside each part.
 *
 * Use {@link OpenXmlPackage.getRelationships} or {@link OpenXmlPart.getRelationships}
 * to retrieve relationships, and {@link RelationshipType} constants for type comparisons.
 *
 * @example
 * ```typescript
 * const rels = await mainPart.getRelationships();
 * for (const rel of rels) {
 *   console.log(`${rel.getId()} [${rel.getType()}] -> ${rel.getTargetFullName()}`);
 * }
 * ```
 */
export class OpenXmlRelationship {
  private pkg: OpenXmlPackage;
  private part: OpenXmlPart | null;
  private id: string;
  private type: string;
  private target: string;
  private targetMode: string | null;

  /** @internal */
  constructor(pkg: OpenXmlPackage, part: OpenXmlPart | null, id: string, type: string, target: string, targetMode: string | null) {
    this.pkg = pkg;
    this.part = part;
    this.id = id;
    this.type = type;
    this.target = target;
    this.targetMode = targetMode;
  }

  /** Returns the {@link OpenXmlPackage} that contains this relationship. */
  getPkg(): OpenXmlPackage {
    return this.pkg;
  }

  /** Returns the source part of this relationship, or `null` for package-level relationships. */
  getPart(): OpenXmlPart | null {
    return this.part;
  }

  /** Returns the unique relationship ID (e.g., `"rId1"`). */
  getId(): string {
    return this.id;
  }

  /** Returns the relationship type URI. Compare with {@link RelationshipType} constants. */
  getType(): string {
    return this.type;
  }

  /** Returns the raw target URI (may be relative). Use {@link getTargetFullName} for a resolved URI. */
  getTarget(): string {
    return this.target;
  }

  /** Returns `"External"` for external targets, or `null` for internal targets. */
  getTargetMode(): string | null {
    return this.targetMode;
  }

  /**
   * Returns the fully resolved target URI.
   *
   * @remarks
   * For internal targets, resolves the relative path against the source part's directory.
   * For external targets and absolute paths, returns the target as-is.
   */
  getTargetFullName(): string {
    if (this.targetMode === "External" || this.target.startsWith("/")) {
      return this.target;
    }
    if (this.part === null) {
      return "/" + this.target;
    }
    const uri = this.part.getUri();
    const dir = uri.substring(0, uri.lastIndexOf("/") + 1);
    return dir + this.target;
  }
}
