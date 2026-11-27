import { FileData, SharePointDocument } from "../Interfaces";
import { escapeXml } from "../utils";

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

/** Adds an "objectid equals this record" condition to a Notes view's FetchXML. */
export function addRegardingFilter(fetchXml: string, recordId: string): string {
  const newCond = `<condition attribute="objectid" operator="eq" value="${escapeXml(recordId)}"/>`;

  if (fetchXml.match(/<filter[^>]*>/i)) {
    return fetchXml.replace(/(<filter[^>]*>)/i, `$1${newCond}`);
  }
  const wrapper = `<filter type="and">${newCond}</filter>`;
  return fetchXml.replace(/<\/entity>/i, `${wrapper}</entity>`);
}

/** Notes whose file name contains the search text. */
export function filterNotes(files: FileData[], searchText: string): FileData[] {
  return files.filter(
    (file) => file.filename && file.filename.toLowerCase().includes(searchText.toLowerCase()),
  );
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
