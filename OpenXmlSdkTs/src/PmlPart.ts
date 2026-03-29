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

export class PmlPart extends OpenXmlPart {
  async commentAuthorsPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.commentAuthors) as Promise<PmlPart | undefined>;
  }

  async handoutMasterPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.handoutMaster) as Promise<PmlPart | undefined>;
  }

  async notesMasterPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.notesMaster) as Promise<PmlPart | undefined>;
  }

  async notesSlidePart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.notesSlide) as Promise<PmlPart | undefined>;
  }

  async presentationPropertiesPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.presentationProperties) as Promise<PmlPart | undefined>;
  }

  async tableStylesPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.tableStyles) as Promise<PmlPart | undefined>;
  }

  async userDefinedTagsPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.userDefinedTags) as Promise<PmlPart | undefined>;
  }

  async viewPropertiesPart(): Promise<PmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.viewProperties) as Promise<PmlPart | undefined>;
  }

  async slideMasterParts(): Promise<PmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.slideMaster) as Promise<PmlPart[]>;
  }

  async slideParts(): Promise<PmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.slide) as Promise<PmlPart[]>;
  }
}
