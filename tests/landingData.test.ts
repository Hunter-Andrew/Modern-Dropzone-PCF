import {
  isActivityRecord,
  loadNotes,
  loadNoteViews,
  readRememberLocation,
  readUserSettings,
  saveUserSettings,
  shouldFilterByLocation,
  waitForSharePointLocation,
} from "../Dropzone/Landing/landingData";
import { NoteView } from "../Dropzone/Interfaces";
import { mockContext, mockFetch, RECORD_ID } from "./mockContext";

beforeEach(() => {
  localStorage.clear();
  mockFetch();
  jest.spyOn(console, "error").mockImplementation(() => undefined);
  jest.spyOn(console, "warn").mockImplementation(() => undefined);
});

afterEach(() => {
  jest.restoreAllMocks();
  jest.useRealTimers();
});

const view = (name: string, fetchxml = "<fetch><entity></entity></fetch>"): NoteView => ({
  savedqueryid: `id-${name}`,
  name,
  fetchxml,
});

describe("loadNoteViews", () => {
  const views = [
    view("My Notes"),
    view("Notes Associated View"),
    view("All Notes"),
    view("Quick Find Annotations"),
  ];

  it("hides the system views and puts All Notes first", async () => {
    const { context } = mockContext({ records: { savedquery: views } });
    const result = await loadNoteViews(context, "");
    expect(result.views.map((v) => v.name)).toEqual(["All Notes", "My Notes"]);
    expect(result.defaultViewId).toBe("id-All Notes");
  });

  it("selects the configured default view, ignoring case", async () => {
    const { context } = mockContext({ records: { savedquery: views } });
    expect((await loadNoteViews(context, "my notes")).defaultViewId).toBe("id-My Notes");
  });

  it("falls back to All Notes when the configured view doesn't exist", async () => {
    const { context } = mockContext({ records: { savedquery: views } });
    expect((await loadNoteViews(context, "Missing")).defaultViewId).toBe("id-All Notes");
  });

  it("returns no views when loading fails", async () => {
    const { context, retrieveMultipleRecords } = mockContext();
    retrieveMultipleRecords.mockRejectedValueOnce(new Error("offline"));
    expect(await loadNoteViews(context, "")).toEqual({ views: [] });
  });
});

describe("loadNotes", () => {
  const row = {
    annotationid: "n1",
    filename: "a.pdf",
    filesize: 10,
    documentbody: "AAA",
    mimetype: "application/pdf",
    createdon: "2025-01-02T00:00:00Z",
    subject: "s",
    notetext: "t",
    objecttypecode: "account",
  };

  it("queries the view's FetchXML with the record filter added", async () => {
    const { context, retrieveMultipleRecords } = mockContext({ records: { annotation: [row] } });
    const notes = await loadNotes(context, view("All Notes"));

    const query = decodeURIComponent(retrieveMultipleRecords.mock.calls[0][1] as string);
    expect(query).toContain(`<condition attribute="objectid" operator="eq" value="${RECORD_ID}"/>`);
    expect(query).toContain("$select=filename,filesize,documentbody,mimetype,annotationid");
    expect(notes).toEqual([
      {
        noteId: "n1",
        filename: "a.pdf",
        filesize: 10,
        documentbody: "AAA",
        mimetype: "application/pdf",
        createdon: new Date("2025-01-02T00:00:00Z"),
        subject: "s",
        notetext: "t",
      },
    ]);
  });

  it("returns no notes for a view without FetchXML", async () => {
    const { context, retrieveMultipleRecords } = mockContext({ records: { annotation: [row] } });
    expect(await loadNotes(context, view("Empty", ""))).toEqual([]);
    expect(retrieveMultipleRecords).not.toHaveBeenCalled();
  });

  it("without a view, loads the record's notes of the record's table", async () => {
    const other = { ...row, annotationid: "n2", objecttypecode: "contact" };
    const { context } = mockContext({ records: { annotation: [row, other] } });
    expect((await loadNotes(context)).map((n) => n.noteId)).toEqual(["n1"]);
  });
});

describe("shouldFilterByLocation", () => {
  const location = (name: string) => ({ name, sharepointdocumentlocationid: name });

  it("is false when the only location is the automatic default one", () => {
    expect(
      shouldFilterByLocation(
        [location("Documents on Default Site 1")],
        "Documents on Default Site 1",
      ),
    ).toBe(false);
  });

  it("is true otherwise", () => {
    expect(shouldFilterByLocation([location("Custom")], "Custom")).toBe(true);
    expect(
      shouldFilterByLocation(
        [location("Documents on Default Site 1"), location("Custom")],
        "Documents on Default Site 1",
      ),
    ).toBe(true);
  });
});

describe("waitForSharePointLocation", () => {
  it("resolves true once a location exists", async () => {
    jest.useFakeTimers();
    const fetchMock = mockFetch({ locations: [] });
    const { context } = mockContext();
    const result = waitForSharePointLocation(context, () => false, 5, 1000);

    await jest.advanceTimersByTimeAsync(1000);
    mockFetch({ locations: [{ name: "Docs", sharepointdocumentlocationid: "loc1" }] });
    await jest.advanceTimersByTimeAsync(1000);

    await expect(result).resolves.toBe(true);
    expect(fetchMock).toHaveBeenCalled();
  });

  it("resolves false when the attempts run out", async () => {
    jest.useFakeTimers();
    mockFetch({ locations: [] });
    const result = waitForSharePointLocation(mockContext().context, () => false, 3, 1000);
    await jest.advanceTimersByTimeAsync(3000);
    await expect(result).resolves.toBe(false);
  });

  it("stops when cancelled", async () => {
    jest.useFakeTimers();
    const fetchMock = mockFetch({ locations: [] });
    const result = waitForSharePointLocation(mockContext().context, () => true, 3, 1000);
    await jest.advanceTimersByTimeAsync(1000);
    await expect(result).resolves.toBe(false);
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe("isActivityRecord", () => {
  it("is true for activity tables", async () => {
    mockFetch({ schemaName: "Email" });
    expect(await isActivityRecord(mockContext().context)).toBe(true);
  });

  it("is false for other tables", async () => {
    mockFetch({ schemaName: "Account" });
    expect(await isActivityRecord(mockContext().context)).toBe(false);
  });
});

describe("user settings", () => {
  it("saves and reads back the selected location", async () => {
    await saveUserSettings(mockContext().context, true, "loc1", "Docs");
    expect(readUserSettings()).toEqual({
      tableName: "account",
      NotesOrSharePoint: true,
      selectedDocumentLocation: "loc1",
      selectedDocumentLocationName: "Docs",
    });
  });

  it("reads nothing when nothing was saved", () => {
    expect(readUserSettings()).toBeNull();
  });

  // Documents current behaviour, which is due to be fixed: the toggle is
  // false with no saved settings and undefined otherwise, because saved
  // settings never include entityId.
  it("readRememberLocation: false with no settings, undefined after saving", async () => {
    const { context } = mockContext();
    expect(await readRememberLocation(context)).toBe(false);
    await saveUserSettings(context, true, "loc1", "Docs");
    expect(await readRememberLocation(context)).toBeUndefined();
  });
});
