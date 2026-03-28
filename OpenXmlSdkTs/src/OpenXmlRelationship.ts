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

export class OpenXmlRelationship {
  private pkg: OpenXmlPackage;
  private part: OpenXmlPart | null;
  private id: string;
  private type: string;
  private target: string;
  private targetMode: string | null;

  constructor(
    pkg: OpenXmlPackage,
    part: OpenXmlPart | null,
    id: string,
    type: string,
    target: string,
    targetMode: string | null,
  ) {
    this.pkg = pkg;
    this.part = part;
    this.id = id;
    this.type = type;
    this.target = target;
    this.targetMode = targetMode;
  }

  getPkg(): OpenXmlPackage {
    return this.pkg;
  }

  getPart(): OpenXmlPart | null {
    return this.part;
  }

  getId(): string {
    return this.id;
  }

  getType(): string {
    return this.type;
  }

  getTarget(): string {
    return this.target;
  }

  getTargetMode(): string | null {
    return this.targetMode;
  }
}
