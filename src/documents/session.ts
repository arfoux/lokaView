import type { OpenedDocument } from "./types";

export class DocumentSession {
  private activeDocument?: OpenedDocument;

  get current(): OpenedDocument | undefined {
    return this.activeDocument;
  }

  setActive(document: OpenedDocument): void {
    this.activeDocument?.dispose();
    this.activeDocument = document;
  }

  clear(): void {
    this.activeDocument?.dispose();
    this.activeDocument = undefined;
  }
}
