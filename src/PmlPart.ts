/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { OpenXmlPart } from "./OpenXmlPart";
import { RelationshipType } from "./RelationshipType";

/**
 * PowerPoint-specific part with convenience methods for navigating presentation structure.
 *
 * @remarks
 * Extends {@link OpenXmlPart} with typed accessors for slides, slide masters,
 * notes, and other PowerPoint-specific parts.
 *
 * @example
 * ```typescript
 * const presentation = await doc.presentationPart();
 * const slides = await presentation!.slideParts();
 * for (const slide of slides) {
 *   const xDoc = await slide.getXDocument();
 *   console.log(slide.getUri());
 * }
 * ```
 *
 * @category Class and Type Reference
 */
export class PmlPart extends OpenXmlPart {
  /** Returns the comment authors part, or `undefined` if not present. */
  async commentAuthorsPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.commentAuthors) as Promise<PmlPart | undefined>;
  }

  /** Returns the handout master part, or `undefined` if not present. */
  async handoutMasterPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.handoutMaster) as Promise<PmlPart | undefined>;
  }

  /** Returns the notes master part, or `undefined` if not present. */
  async notesMasterPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.notesMaster) as Promise<PmlPart | undefined>;
  }

  /** Returns the notes slide part, or `undefined` if not present. */
  async notesSlidePart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.notesSlide) as Promise<PmlPart | undefined>;
  }

  /** Returns the presentation properties part, or `undefined` if not present. */
  async presentationPropertiesPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.presentationProperties) as Promise<PmlPart | undefined>;
  }

  /** Returns the table styles part, or `undefined` if not present. */
  async tableStylesPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.tableStyles) as Promise<PmlPart | undefined>;
  }

  /** Returns the user-defined tags part, or `undefined` if not present. */
  async userDefinedTagsPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.userDefinedTags) as Promise<PmlPart | undefined>;
  }

  /** Returns the view properties part, or `undefined` if not present. */
  async viewPropertiesPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.viewProperties) as Promise<PmlPart | undefined>;
  }

  /** Returns all slide master parts. */
  async slideMasterParts(): Promise<PmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.slideMaster) as Promise<PmlPart[]>;
  }

  /** Returns all slide parts. */
  async slideParts(): Promise<PmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.slide) as Promise<PmlPart[]>;
  }
}
