import { DataverseError, toDataverseError } from "../Dropzone/DataverseActions";
import { isLikelyTooLarge, LIKELY_TOO_LARGE_BYTES } from "../Dropzone/Landing/uploadErrors";

const MB = 1024 * 1024;

describe("isLikelyTooLarge", () => {
  it("is true for HTTP 413", () => {
    expect(
      isLikelyTooLarge(new DataverseError("Failed to upload document (HTTP 413)", 413), MB),
    ).toBe(true);
  });

  it("is true when the message is about size", () => {
    for (const message of [
      "Maximum request length exceeded.",
      "The file is too large",
      "Request content length exceeds the limit",
    ]) {
      expect(isLikelyTooLarge(new DataverseError(message, 400), MB)).toBe(true);
    }
  });

  it("treats any failure of a large file as too large", () => {
    expect(isLikelyTooLarge(new TypeError("Failed to fetch"), LIKELY_TOO_LARGE_BYTES)).toBe(true);
  });

  it("is false for other failures of smaller files", () => {
    expect(isLikelyTooLarge(new DataverseError("Access is denied", 403), 10 * MB)).toBe(false);
    expect(isLikelyTooLarge(new TypeError("Failed to fetch"), 10 * MB)).toBe(false);
  });
});

describe("toDataverseError", () => {
  const response = (status: number, body: string, statusText = "") =>
    ({ status, statusText, text: async () => body }) as Response;

  it("uses Dataverse's error message from the body", async () => {
    const error = await toDataverseError(
      response(400, JSON.stringify({ error: { code: "0x0", message: "Folder not found" } })),
      "Failed to upload document",
    );
    expect(error).toBeInstanceOf(DataverseError);
    expect(error.status).toBe(400);
    expect(error.message).toBe("Failed to upload document (HTTP 400): Folder not found");
  });

  it("falls back to the raw body, then the status text", async () => {
    expect(
      (await toDataverseError(response(413, "Request Entity Too Large"), "Upload")).message,
    ).toBe("Upload (HTTP 413): Request Entity Too Large");
    expect((await toDataverseError(response(500, "", "Server Error"), "Upload")).message).toBe(
      "Upload (HTTP 500): Server Error",
    );
    expect((await toDataverseError(response(500, ""), "Upload")).message).toBe("Upload (HTTP 500)");
  });
});
