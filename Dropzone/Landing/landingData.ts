import { IInputs } from "../generated/ManifestTypes";
import {
  getNoteViews,
  getRelatedNotes,
  getSharePointFolderData,
  getSharePointLocations,
} from "../DataverseActions";
import { DocumentLocation, FileData, NoteView, SharePointDocument } from "../Interfaces";
import { getEntityMetadata, isActivityType } from "../utils";
import { addRegardingFilter } from "./landingHelpers";

// Data loading for Landing: reading notes, note views, SharePoint locations
// and folders, and the user's saved settings. Nothing here touches React
// state; Landing calls these and stores the results.

type Context = ComponentFramework.Context<IInputs>;

const NOTE_COLUMNS =
  "filename,filesize,documentbody,mimetype,annotationid,createdon,subject,notetext";

const EXCLUDED_NOTE_VIEWS = new Set<string>([
  "Notes Associated View",
  "Notes",
  "Notes Advanced Find View",
  "Notes Lookup View",
  "Quick Find Annotations",
]);

/** The location Dynamics creates automatically for a record. */
export const DEFAULT_LOCATION_NAME = "Documents on Default Site 1";

const USER_SETTINGS_KEY = "userSettings";

export const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

// ---------------------------------------------------------------------------
// Notes
// ---------------------------------------------------------------------------

export interface NoteViewsResult {
  views: NoteView[];
  /** The view to select first: the configured default, else "All Notes", else the first. */
  defaultViewId?: string;
}

/**
 * Loads the note views the user can pick from, with "All Notes" first.
 * Returns no views (and logs why) if they can't be loaded.
 */
export async function loadNoteViews(
  context: Context,
  defaultViewName: string
): Promise<NoteViewsResult> {
  try {
    const res = await getNoteViews(context);
    if (!res.success) {
      console.warn(res.message);
      return { views: [] };
    }

    let views = res.data.filter((v) => !EXCLUDED_NOTE_VIEWS.has(v.name));
    const allNotes = views.find((v) => v.name.toLowerCase() === "all notes");
    if (allNotes) {
      views = [allNotes, ...views.filter((v) => v !== allNotes)];
    }

    const findByName = (name: string) =>
      views.find((v) => v.name.toLowerCase() === name.toLowerCase());
    const defaultView =
      (defaultViewName && findByName(defaultViewName)) ||
      findByName("All Notes") ||
      views[0];

    return { views, defaultViewId: defaultView?.savedqueryid };
  } catch (err) {
    console.error("Failed to load note views:", err);
    return { views: [] };
  }
}

/**
 * Loads the notes for the current record. With a view, its FetchXML is used
 * with a regarding filter added; without one, all notes for the record.
 */
export async function loadNotes(context: Context, view?: NoteView): Promise<FileData[]> {
  let rows: any[] = [];

  if (view) {
    if (view.fetchxml) {
      const meta = await getEntityMetadata(context);
      const recordId = meta?.entityId?.replace(/[{}]/g, "");
      const fetch = recordId ? addRegardingFilter(view.fetchxml, recordId) : view.fetchxml;
      const res = await context.webAPI.retrieveMultipleRecords(
        "annotation",
        `?fetchXml=${encodeURIComponent(fetch)}&$select=${NOTE_COLUMNS}`
      );
      rows = res.entities;
    }
  } else {
    const res = await getRelatedNotes(context);
    if (res.success) rows = res.data;
  }

  return rows.map((r) => ({
    filename: r.filename,
    filesize: r.filesize,
    documentbody: r.documentbody,
    mimetype: r.mimetype,
    noteId: r.annotationid,
    createdon: new Date(r.createdon),
    subject: r.subject,
    notetext: r.notetext,
  }));
}

// ---------------------------------------------------------------------------
// SharePoint
// ---------------------------------------------------------------------------

/** The SharePoint document locations of the current record. */
export async function loadSharePointLocations(context: Context): Promise<DocumentLocation[]> {
  return getSharePointLocations(context);
}

/** True when the record has no SharePoint document location yet. */
export async function hasNoSharePointLocation(context: Context): Promise<boolean> {
  const locations = await getSharePointLocations(context);
  return locations.length === 0;
}

/**
 * False when the record's only location is the one Dynamics created
 * automatically, true otherwise. DataverseActions filters documents by the
 * selected location only when this is true.
 */
export function shouldFilterByLocation(
  locations: DocumentLocation[],
  selectedLocationName: string
): boolean {
  return !(locations.length === 1 && selectedLocationName === DEFAULT_LOCATION_NAME);
}

/** The documents and folders in one folder of a SharePoint location. */
export async function loadSharePointFolder(
  context: Context,
  folderPath: string,
  locationId: string | null,
  locationName: string,
  filterByLocation?: boolean
): Promise<SharePointDocument[]> {
  return getSharePointFolderData(context, folderPath, locationId, locationName, filterByLocation);
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
  intervalMs = 3000
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
// Record
// ---------------------------------------------------------------------------

/** True when the current record is an activity (email, task, ...). */
export async function isActivityRecord(context: Context): Promise<boolean> {
  const metadata = await getEntityMetadata(context);
  return !!metadata && isActivityType(metadata.schemaName);
}

// ---------------------------------------------------------------------------
// User settings (browser localStorage)
// ---------------------------------------------------------------------------

export interface UserSettings {
  tableName?: string;
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
  if (
    settings.tableName !== metadata?.schemaName ||
    settings.entityId !== metadata?.entityId
  ) {
    return undefined;
  }
  return typeof settings.NotesOrSharePoint === "undefined"
    ? settings.NotesOrSharePoint
    : undefined;
}

export async function saveUserSettings(
  context: Context,
  notesOrSharePoint: boolean,
  selectedDocumentLocation: string | null,
  selectedDocumentLocationName: string
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  const settings: UserSettings = {
    tableName: metadata?.schemaName,
    NotesOrSharePoint: notesOrSharePoint,
    selectedDocumentLocation,
    selectedDocumentLocationName,
  };
  localStorage.setItem(USER_SETTINGS_KEY, JSON.stringify(settings));
}
