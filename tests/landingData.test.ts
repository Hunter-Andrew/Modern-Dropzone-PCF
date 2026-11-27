import {
  loadLocationUrl,
  readRememberLocation,
  readUserSettings,
  saveUserSettings,
  shouldFilterByLocation,
  waitForSharePointLocation,
} from "../Dropzone/Landing/landingData";
import { mockContext, mockFetch } from "./mockContext";

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

describe("loadLocationUrl", () => {
  it("asks Dataverse for the location's folder URL", async () => {
    const url = "https://contoso.sharepoint.com/sites/crm/account/Contoso_1234";
    const fetchMock = mockFetch({ locationUrl: url });
    expect(await loadLocationUrl(mockContext().context, "{LOC-1}")).toBe(url);
    expect(fetchMock.mock.calls.map((call) => call[0])).toContain(
      "https://org.example/api/data/v9.0/sharepointdocumentlocations(LOC-1)/Microsoft.Dynamics.CRM.RetrieveAbsoluteAndSiteCollectionUrl()",
    );
  });

  it('returns "" when the lookup fails', async () => {
    mockFetch();
    const failing = jest.fn(async (url: string) =>
      url.includes("EntityDefinitions")
        ? ({
            ok: true,
            status: 200,
            json: async () => ({ SchemaName: "Account", LogicalCollectionName: "accounts" }),
          } as Response)
        : ({ ok: false, status: 404, statusText: "", text: async () => "" } as Response),
    );
    global.fetch = failing as unknown as typeof fetch;
    expect(await loadLocationUrl(mockContext().context, "loc1")).toBe("");
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
