import { IInputs } from "../Dropzone/generated/ManifestTypes";

// A minimal stand-in for the PCF context and the Dataverse Web API, enough
// for the data functions under test.

export const RECORD_ID = "11111111-2222-3333-4444-555555555555";

export interface MockOptions {
  entityId?: string;
  entityTypeName?: string;
  /** Rows returned by webAPI.retrieveMultipleRecords, per entity name. */
  records?: Record<string, unknown[]>;
  parameters?: Record<string, unknown>;
}

export function mockContext(options: MockOptions = {}) {
  const retrieveMultipleRecords = jest.fn(async (entity: string, _query?: string) => ({
    entities: options.records?.[entity] ?? [],
  }));
  const parameters: Record<string, { raw: unknown }> = {};
  for (const [name, value] of Object.entries(options.parameters ?? {})) {
    parameters[name] = { raw: value };
  }

  const context = {
    page: {
      entityId: options.entityId ?? RECORD_ID,
      entityTypeName: options.entityTypeName ?? "account",
      getClientUrl: () => "https://org.example",
    },
    parameters,
    webAPI: { retrieveMultipleRecords },
    resources: { getString: (key: string) => `[${key}]` },
  };
  return {
    context: context as unknown as ComponentFramework.Context<IInputs>,
    retrieveMultipleRecords,
  };
}

/**
 * Makes fetch answer the entity metadata request (used by getEntityMetadata)
 * and, optionally, SharePoint document locations.
 */
export function mockFetch(options: { schemaName?: string; locations?: unknown[] } = {}) {
  const fetchMock = jest.fn(async (url: string) => {
    const body = url.includes("EntityDefinitions")
      ? { SchemaName: options.schemaName ?? "Account", LogicalCollectionName: "accounts" }
      : url.includes("sharepointdocumentlocations")
        ? { value: options.locations ?? [] }
        : {};
    return { ok: true, status: 200, json: async () => body } as Response;
  });
  global.fetch = fetchMock as unknown as typeof fetch;
  return fetchMock;
}
