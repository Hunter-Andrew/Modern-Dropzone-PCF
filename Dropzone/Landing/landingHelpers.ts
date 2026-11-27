import { SharePointDocument } from "../Interfaces";

/** Formats a byte count as "1.23 MB", "4.56 KB" or "789 Bytes". */
export function formatFileSize(sizeInBytes: number): string {
  const sizeInKB = sizeInBytes / 1024;
  const sizeInMB = sizeInBytes / 1048576;

  if (sizeInMB >= 1) {
    return `${sizeInMB.toFixed(2)} MB`;
  } else if (sizeInKB >= 1) {
    return `${sizeInKB.toFixed(2)} KB`;
  } else {
    return `${sizeInBytes} Bytes`;
  }
}

/**
 * Shortens a long name for a file tile. Files keep the start of the name, the
 * last few characters and the extension ("Quarterly...ion.pdf"); folders, and
 * names without an extension, are cut at the end ("Invoices & Rece...").
 */
export function middleEllipsis(filename: string, maxLength = 18, isFolder = false): string {
  if (filename.length < maxLength) {
    return filename;
  }

  const lastDotIndex = filename.lastIndexOf(".");
  if (isFolder || lastDotIndex === -1) {
    return `${filename.substring(0, maxLength - 3)}...`;
  }

  const extension = filename.substring(lastDotIndex + 1);
  const name = filename.substring(0, lastDotIndex);
  const startChars = 9;
  const endChars = 3;
  const start = name.substring(0, startChars);
  const end = name.substring(name.length - endChars);

  if (end.length < 3 && name.length - 3 > startChars) {
    return `${start}...${name.slice(-3)}.${extension}`;
  }
  return `${start}...${end}.${extension}`;
}

/** Lower-case extension used to pick a file-type icon ("folder" when there's no name). */
export function getFileExtension(filename?: string): string {
  if (!filename) {
    return "folder";
  }
  const extension = filename.split(".").pop()?.toLowerCase();
  return extension ? extension : "txt";
}

/** SharePoint items matching the search text, folders first, then by name. */
export function filterAndSortSharePointItems(
  items: SharePointDocument[],
  searchText: string,
): SharePointDocument[] {
  return items
    .filter((item) => item.fullname.toLowerCase().includes(searchText.toLowerCase().trim()))
    .sort((a, b) => {
      if (a.filetype === "folder" && b.filetype !== "folder") {
        return -1;
      } else if (a.filetype !== "folder" && b.filetype === "folder") {
        return 1;
      }
      return a.fullname.toLowerCase().localeCompare(b.fullname.toLowerCase());
    });
}

/**
 * The SharePoint URL of the folder being shown, or "" if it can't be worked out.
 *
 * The parent of any item's URL is exactly this folder, so that is used when
 * the folder has items. An empty folder falls back to the location's URL plus
 * the open subfolder's path.
 */
export function currentFolderUrl(
  locationUrl: string,
  folderPath: string,
  items: SharePointDocument[],
): string {
  const itemUrl = items.find((item) => item.absoluteurl)?.absoluteurl;
  if (itemUrl) {
    return itemUrl.replace(/\/+$/, "").replace(/\/[^/]*$/, "");
  }
  if (!locationUrl) return "";
  const base = locationUrl.replace(/\/+$/, "");
  if (!folderPath) return base;
  const path = folderPath
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");
  return `${base}/${path}`;
}
