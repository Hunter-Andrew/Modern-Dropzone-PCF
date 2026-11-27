import { DataverseError } from "../DataverseActions";
import { errorMessage } from "../utils";

// Deciding whether a failed upload was most likely rejected for its size.
//
// Uploads go through Dataverse's UploadDocument action (the same one the
// Documents grid uses), which takes the whole file as base64 in one request.
// Dataverse doesn't document a limit for it; community reports put it around
// 50 MB. Over that, a failure is treated as a size problem whatever its
// message, since large requests can also fail as a dropped connection.

/** Failed uploads at or above this size are treated as too large. */
export const LIKELY_TOO_LARGE_BYTES = 50 * 1024 * 1024;

const SIZE_MESSAGE =
  /too large|too big|exceed|maximum (allowed|request|file|size)|size limit|request length|content length|payload/i;

/** True when an upload of a file this size most likely failed because of its size. */
export function isLikelyTooLarge(error: unknown, fileSize: number): boolean {
  if (error instanceof DataverseError && error.status === 413) return true;
  if (SIZE_MESSAGE.test(errorMessage(error))) return true;
  return fileSize >= LIKELY_TOO_LARGE_BYTES;
}
