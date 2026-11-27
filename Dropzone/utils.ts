import { LocalStrings } from "./consts/LocalStrings";
import { IInputs } from "./generated/ManifestTypes";
import { EntityMetadata } from "./Interfaces";
import { ActivityType } from "./Interfaces";

type Context = ComponentFramework.Context<IInputs>;

/**
 * The parts of `context.page` this control uses. Model-driven apps provide
 * `page`, but the PCF type definitions don't declare it.
 */
export interface PageInfo {
  entityId: string;
  entityTypeName: string;
  getClientUrl(): string;
}

export function getPage(context: Context | undefined): PageInfo | undefined {
  return (context as unknown as { page?: PageInfo } | undefined)?.page;
}

/** The id of the record the form shows, or "" when it hasn't been saved yet. */
export function getRecordId(context: Context | undefined): string {
  return getPage(context)?.entityId ?? "";
}

/** The message of a caught error, whatever was thrown. */
export function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export async function getEntityMetadata(
  context: ComponentFramework.Context<IInputs>
): Promise<EntityMetadata | null> {
  const page = getPage(context);
  if (!page) {
    console.warn("Component Framework context is not available. (utils)");
    return null;
  }

  const dynamicsUrl = page.getClientUrl();
  if (!dynamicsUrl) {
    console.error("Unable to retrieve client URL.");
    return null;
  }

  const entityName = page.entityTypeName;
  if (!entityName) {
    console.error("Unable to retrieve entity type name.");
    return null;
  }
  const apiUrl = `${dynamicsUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${entityName}')?$select=SchemaName,LogicalCollectionName`;

  const entityId = page.entityId;
  if (!entityId) {
    console.error("Unable to retrieve entity ID.");
    return null;
  }

  try {
    const response = await fetch(apiUrl, {
      headers: {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        Accept: "application/json",
        "Content-Type": "application/json; charset=utf-8",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return {
      schemaName: data.SchemaName.toLowerCase(),
      logicalCollectionName: data.LogicalCollectionName,
      clientUrl: dynamicsUrl,
      entityId: entityId,
    };
  } catch (error) {
    console.error("Error fetching entity metadata:", error);
    return null;
  }
}

/**
 * Escapes a value for use inside a FetchXML attribute or text node.
 * Without this, names containing characters such as & " ' < > produce invalid
 * FetchXML and the query fails (or matches the wrong records).
 */
export function escapeXml(value: string | null | undefined): string {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// Characters SharePoint does not allow in file or folder names.
const SHAREPOINT_INVALID_CHARS = /["*:<>?/\\|]/;

/** Returns true if the name can be used for a SharePoint folder. */
export function isValidSharePointFolderName(name: string): boolean {
  return (
    name.trim().length > 0 &&
    !SHAREPOINT_INVALID_CHARS.test(name) &&
    !/^[\s.]|[\s.]$/.test(name)
  );
}

export function isPDF(mime: string) {
  return mime === "application/pdf";
}

export function isImage(mime: string) {
  return mime.startsWith("image/");
}

export function createDataUri(mimetype: string, base64: string): string {
  return `data:${mimetype};base64,${base64}`;
}

export function isActivityType(schemaName: string): schemaName is ActivityType {
  return [
    "email",
    "phonecall",
    "appointment",
    "task",
    "fax",
    "letter",
    "serviceappointment",
    "campaignresponse",
    "campaignactivity",
    "bulkoperation",
    "socialactivity",
    "recurringappointmentmaster",
    "appointmentrecurrence",
  ].includes(schemaName);
}
export function getControlValue(context: Context, parameter: keyof IInputs) {
  const property = context.parameters[parameter];
  return (property && "raw" in property ? property.raw : null) || "";
}
type Values<T> = T extends string ? T : { [K in keyof T]: Values<T[K]> }[keyof T];
export type LocalizationKey = Values<typeof LocalStrings>;
/** Looks up a localized string; passed to child components as `t`. */
export type Translate = (key: LocalizationKey) => string;

export const getLocalString = (
  context: Context,
  key: LocalizationKey
): string => context.resources.getString(key);

export function b64toBlob(b64Data: string, contentType = '', sliceSize = 512) {
  if (b64Data.startsWith("data:application/pdf;base64,")) {
    b64Data = b64Data.replace("data:application/pdf;base64,", "");
  }

  const byteCharacters = atob(b64Data);
  const byteArrays = [];

  for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
    const slice = byteCharacters.slice(offset, offset + sliceSize);
    const byteNumbers = new Array(slice.length);
    for (let i = 0; i < slice.length; i++) {
      byteNumbers[i] = slice.charCodeAt(i);
    }
    byteArrays.push(new Uint8Array(byteNumbers));
  }

  return new Blob(byteArrays, { type: contentType });
}

