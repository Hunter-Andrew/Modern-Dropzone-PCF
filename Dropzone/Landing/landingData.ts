import { IInputs } from "../generated/ManifestTypes";
import {
  getLocationUrl,
  getSharePointFolderData,
  getSharePointLocations,
} from "../DataverseActions";
import { DocumentLocation, SharePointDocument } from "../Interfaces";
import { getEntityMetadata } from "../utils";

// Data loading for Landing: SharePoint locations, folders and folder URLs,
// and the user's saved settings. Nothing here touches React state; Landing
// calls these and stores the results.

type Context = ComponentFramework.Context<IInputs>;

/** The location Dynamics creates automatically for a record. */
export const DEFAULT_LOCATION_NAME = "Documents on Default Site 1";

const USER_SETTINGS_KEY = "userSettings";

export const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

// ---------------------------------------------------------------------------
// SharePoint
// ---------------------------------------------------------------------------

/** The SharePoint document locations of the current record. */
export async function loadSharePointLocations(context: Context): Promise<DocumentLocation[]> {
  return getSharePointLocations(context);
}

/**
 * False when the record's only location is the one Dynamics created
 * automatically, true otherwise. DataverseActions filters documents by the
 * selected location only when this is true.
 */
export function shouldFilterByLocation(
  locations: DocumentLocation[],
  selectedLocationName: string,
): boolean {
  return !(locations.length === 1 && selectedLocationName === DEFAULT_LOCATION_NAME);
}

/** The documents and folders in one folder of a SharePoint location. */
export async function loadSharePointFolder(
  context: Context,
  folderPath: string,
  locationId: string | null,
  locationName: string,
  filterByLocation?: boolean,
): Promise<SharePointDocument[]> {
  return getSharePointFolderData(context, folderPath, locationId, locationName, filterByLocation);
}

/**
 * The URL of a location's folder in SharePoint, or "" if Dataverse can't
 * provide it (see currentFolderUrl for the fallback).
 */
export async function loadLocationUrl(context: Context, locationId: string): Promise<string> {
  try {
    return await getLocationUrl(context, locationId);
  } catch (err) {
    console.warn("Getting the SharePoint folder URL failed:", err);
    return "";
  }
}

/**
 * After a record's first save, Dynamics creates its SharePoint folder in the
 * background (e.g. when a Documents grid on the form loads). Polls until a
 * document location exists. Resolves true once one is found, false when the
 * attempts run out or `isCancelled` returns true.
 */
export async function waitForSharePointLocation(
  context: Context,
  isCancelled: () => boolean,
  attempts = 8,
  intervalMs = 3000,
): Promise<boolean> {
  for (let i = 0; i < attempts; i++) {
    await delay(intervalMs);
    if (isCancelled()) return false;
    try {
      const locations = await getSharePointLocations(context);
      if (locations.length > 0) return true;
    } catch (err) {
      console.warn("Checking for SharePoint document location failed:", err);
    }
  }
  return false;
}

// ---------------------------------------------------------------------------
// User settings (browser localStorage)
// ---------------------------------------------------------------------------

export interface UserSettings {
  tableName?: string;
  /** The "Remember Location?" toggle. The name dates from the old Notes/SharePoint switch. */
  NotesOrSharePoint?: boolean;
  selectedDocumentLocation?: string | null;
  selectedDocumentLocationName?: string;
  entityId?: string;
}

export function readUserSettings(): UserSettings | null {
  const settings = localStorage.getItem(USER_SETTINGS_KEY);
  return settings ? JSON.parse(settings) : null;
}

/**
 * The initial value of the "Remember Location?" toggle. Kept exactly as the
 * control has always worked: false with no saved settings, otherwise
 * undefined in practice, because the saved settings never include entityId.
 */
export async function readRememberLocation(context: Context): Promise<boolean | undefined> {
  const settings = readUserSettings();
  if (!settings) return false;

  const metadata = await getEntityMetadata(context);
  if (settings.tableName !== metadata?.schemaName || settings.entityId !== metadata?.entityId) {
    return undefined;
  }
  return typeof settings.NotesOrSharePoint === "undefined" ? settings.NotesOrSharePoint : undefined;
}

export async function saveUserSettings(
  context: Context,
  rememberLocation: boolean,
  selectedDocumentLocation: string | null,
  selectedDocumentLocationName: string,
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  const settings: UserSettings = {
    tableName: metadata?.schemaName,
    NotesOrSharePoint: rememberLocation,
    selectedDocumentLocation,
    selectedDocumentLocationName,
  };
  localStorage.setItem(USER_SETTINGS_KEY, JSON.stringify(settings));
}
