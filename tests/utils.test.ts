import {
  createDataUri,
  errorMessage,
  escapeXml,
  getControlValue,
  getRecordId,
  isActivityType,
  isImage,
  isPDF,
  isValidSharePointFolderName,
} from "../Dropzone/utils";
import { mockContext, RECORD_ID } from "./mockContext";

describe("escapeXml", () => {
  it("escapes the five XML special characters", () => {
    expect(escapeXml(`Tom & "Jerry's" <files>`)).toBe(
      "Tom &amp; &quot;Jerry&apos;s&quot; &lt;files&gt;",
    );
  });

  it("turns null and undefined into an empty string", () => {
    expect(escapeXml(null)).toBe("");
    expect(escapeXml(undefined)).toBe("");
  });
});

describe("isValidSharePointFolderName", () => {
  it.each(["Contracts", "Invoices & Receipts", "2024 Q1", "a.b"])("accepts %p", (name) => {
    expect(isValidSharePointFolderName(name)).toBe(true);
  });

  it.each([
    "",
    "   ",
    "a/b",
    "a\\b",
    'a"b',
    "a*b",
    "a:b",
    "a<b",
    "a>b",
    "a?b",
    "a|b",
    " lead",
    "trail ",
    ".hidden",
    "dot.",
  ])("rejects %p", (name) => {
    expect(isValidSharePointFolderName(name)).toBe(false);
  });
});

describe("file type checks", () => {
  it("isPDF", () => {
    expect(isPDF("application/pdf")).toBe(true);
    expect(isPDF("image/png")).toBe(false);
  });

  it("isImage", () => {
    expect(isImage("image/png")).toBe(true);
    expect(isImage("application/pdf")).toBe(false);
  });

  it("createDataUri", () => {
    expect(createDataUri("image/png", "AAA")).toBe("data:image/png;base64,AAA");
  });

  it("isActivityType", () => {
    expect(isActivityType("email")).toBe(true);
    expect(isActivityType("account")).toBe(false);
  });
});

describe("getRecordId", () => {
  it("returns the record id from context.page", () => {
    expect(getRecordId(mockContext().context)).toBe(RECORD_ID);
  });

  it('returns "" for an unsaved record or a missing context', () => {
    expect(getRecordId(mockContext({ entityId: "" }).context)).toBe("");
    expect(getRecordId(undefined)).toBe("");
  });
});

describe("getControlValue", () => {
  it("returns the property's value, or an empty string when unset", () => {
    const { context } = mockContext({ parameters: { allowNoteDrops: true } });
    expect(getControlValue(context, "allowNoteDrops")).toBe(true);
    expect(getControlValue(context, "allowSharePointDrops")).toBe("");
  });
});

describe("errorMessage", () => {
  it("reads Error messages and stringifies anything else", () => {
    expect(errorMessage(new Error("boom"))).toBe("boom");
    expect(errorMessage("plain")).toBe("plain");
  });
});
