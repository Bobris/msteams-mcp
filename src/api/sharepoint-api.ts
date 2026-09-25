/**
 * SharePoint/OneDrive file upload API via Microsoft Graph.
 *
 * Teams file attachments require uploading to OneDrive, resolving the file's
 * SharePoint GUID, and granting chat recipients read access before sending
 * a chat message with a `files` property referencing the uploaded file.
 *
 * We use the Graph API (`graph.microsoft.com`) rather than the SharePoint REST
 * API directly because:
 * - The Graph token is already in the MSAL cache (broad delegated permissions)
 * - No need to construct tenant-specific SharePoint URLs manually
 * - The Graph response includes all SharePoint URLs needed for the `files` property
 *
 * Reverse-engineered from Teams web client network interception (2026-07-28).
 */

import { open, type FileHandle } from 'node:fs/promises';
import { basename, extname } from 'node:path';
import { httpRequest } from '../utils/http.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import { getValidGraphToken } from '../auth/token-extractor.js';
import { UPLOAD_CHUNK_BYTES, UPLOAD_READ_BYTES } from '../constants.js';
import { grantAttachmentAccess } from './attachment-sharing.js';
import { GRAPH_FILES_API } from '../utils/api-config.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Subset of the Graph API DriveItem response that we need. */
export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  sharepointIds?: {
    siteId?: string;
    siteUrl?: string;
    webId?: string;
    listId?: string;
    listItemUniqueId?: string;
  };
  parentReference?: {
    driveId?: string;
    sharepointIds?: {
      siteId?: string;
      siteUrl?: string;
      webId?: string;
      listId?: string;
      listItemUniqueId?: string;
    };
  };
}

/** Result of uploading a file. */
export interface UploadFileResult {
  /** The DriveItem ID from Graph/SharePoint. */
  itemId: string;
  /** The file name (may differ from input if renamed due to conflict). */
  fileName: string;
  /** File extension without the dot (e.g., "pdf"). */
  fileType: string;
  /** File size in bytes. */
  fileSize?: number;
  /** The SharePoint personal site base URL (e.g., "https://tenant-my.sharepoint.com/personal/user_domain_com/"). */
  baseUrl: string;
  /** The full SharePoint URL to the file. */
  objectUrl: string;
  /** The web URL from Graph (may differ slightly from objectUrl). */
  webUrl?: string;
  /** SharePoint list item unique ID (used in the `files` property). */
  listItemUniqueId?: string;
  /** The JSON-encoded `files` property string for the chatsvc message body. */
  filesProperty: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

/** Graph API base URL. */
const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

/** The OneDrive folder Teams uses for chat file attachments. */
const TEAMS_CHAT_FILES_FOLDER = 'Microsoft Teams Chat Files';

/** Session responses between chunks (the final response is a DriveItem). */
interface UploadSession {
  uploadUrl?: string;
  nextExpectedRanges?: string[];
}

/** Stream a bounded range using small reads, keeping the file handle open across chunks. */
function fileRange(file: FileHandle, start: number, end: number): ReadableStream<Uint8Array> {
  let position = start;
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      if (position > end) {
        controller.close();
        return;
      }
      const buffer = new Uint8Array(Math.min(UPLOAD_READ_BYTES, end - position + 1));
      const { bytesRead } = await file.read(buffer, 0, buffer.length, position);
      if (!bytesRead) throw new Error('Local file was truncated during upload');
      position += bytesRead;
      controller.enqueue(buffer.subarray(0, bytesRead));
    },
  }, new ByteLengthQueuingStrategy({ highWaterMark: UPLOAD_READ_BYTES }));
}

/** Upload sequential Graph fragments without buffering the file or a whole fragment. */
async function uploadContent(file: FileHandle, fileName: string, size: number, graphToken: string): Promise<Result<DriveItem>> {
  const itemUrl = `${GRAPH_BASE_URL}/me/drive/root:/${encodeURIComponent(TEAMS_CHAT_FILES_FOLDER)}/${encodeURIComponent(fileName)}`;
  const authorization = { Authorization: `Bearer ${graphToken}` };
  // Graph upload sessions need a non-empty byte range; empty files use simple upload.
  if (size === 0) {
    const response = await httpRequest<DriveItem>(`${itemUrl}:/content`, {
      method: 'PUT', headers: { ...authorization, 'Content-Type': 'application/octet-stream' },
      body: new Uint8Array(0), maxRetries: 1, redirect: 'error',
    });
    return response.ok ? ok(response.value.data) : response;
  }
  const session = await httpRequest<UploadSession>(`${itemUrl}:/createUploadSession`, {
    method: 'POST',
    headers: { ...authorization, 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'rename' } }),
    maxRetries: 1, redirect: 'error',
  });
  if (!session.ok) return session;
  let uploadUrl: string;
  try {
    const parsed = new URL(session.value.data.uploadUrl ?? '');
    if (parsed.protocol !== 'https:') throw new Error('Expected HTTPS');
    uploadUrl = parsed.href;
  } catch {
    return err(createError(ErrorCode.API_ERROR, 'Graph returned no valid HTTPS upload session URL'));
  }
  let complete = false;
  try {
    for (let start = 0; start < size; start += UPLOAD_CHUNK_BYTES) {
      const end = Math.min(start + UPLOAD_CHUNK_BYTES, size) - 1;
      const body = fileRange(file, start, end);
      const response = await httpRequest<DriveItem & UploadSession>(uploadUrl, {
        method: 'PUT',
        // The upload URL is preauthenticated. Never send the Graph token here.
        headers: {
          'Content-Type': 'application/octet-stream',
          'Content-Length': String(end - start + 1),
          'Content-Range': `bytes ${start}-${end}/${size}`,
        },
        body, duplex: 'half', redirect: 'error', maxRetries: 1,
      });
      // Cancel unconsumed bytes if the server rejected the request early.
      await body.cancel().catch(() => {});
      if (!response.ok) return response;
      const { status, data } = response.value;
      if (status === 200 || status === 201) {
        if (end !== size - 1 || !data.id || !data.name || data.size !== size) {
          return err(createError(ErrorCode.API_ERROR, 'Upload completed with unexpected file metadata or size'));
        }
        complete = true;
        return ok(data);
      }
      // SharePoint may return either "start-" or "start-end" for the remaining range.
      const nextRange = data.nextExpectedRanges?.[0]?.match(/^(\d+)-(\d*)$/);
      if (status !== 202 || end === size - 1 || !nextRange ||
          Number(nextRange[1]) !== end + 1 || (nextRange[2] && Number(nextRange[2]) !== size - 1)) {
        return err(createError(ErrorCode.API_ERROR, 'Upload session returned an unexpected next byte range'));
      }
    }
    return err(createError(ErrorCode.API_ERROR, 'Upload session did not return a completed file'));
  } finally {
    if (!complete) {
      // Best effort cancellation removes temporary fragments, not completed drive items.
      await httpRequest(uploadUrl, { method: 'DELETE', maxRetries: 1, redirect: 'error' });
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the file extension without the leading dot (e.g., "pdf" from "doc.pdf").
 */
function getFileExtension(fileName: string): string {
  const ext = extname(fileName).toLowerCase().replace(/^\./, '');
  return ext || 'file';
}

/**
 * Extracts the SharePoint personal site base URL from a DriveItem response.
 *
 * The base URL looks like: `https://{tenant}-my.sharepoint.com/personal/{user_folder}/`
 */
function extractBaseUrl(driveItem: DriveItem): string | null {
  // Try sharepointIds.siteUrl first
  const siteUrl = driveItem.sharepointIds?.siteUrl
    ?? driveItem.parentReference?.sharepointIds?.siteUrl;
  if (siteUrl) {
    return siteUrl.endsWith('/') ? siteUrl : `${siteUrl}/`;
  }

  // Fallback: parse from webUrl
  if (driveItem.webUrl) {
    try {
      const url = new URL(driveItem.webUrl);
      // webUrl is like: https://{tenant}-my.sharepoint.com/personal/{user}/Documents/Microsoft Teams Chat Files/{file}
      const pathParts = url.pathname.split('/');
      // Find "personal" in the path and take the next segment
      const personalIndex = pathParts.indexOf('personal');
      if (personalIndex >= 0 && pathParts[personalIndex + 1]) {
        const base = `${url.protocol}//${url.host}/personal/${pathParts[personalIndex + 1]}/`;
        return base;
      }
    } catch {
      // Ignore parse errors
    }
  }

  return null;
}

/**
 * Builds the `files` property JSON string from a Graph API DriveItem response.
 *
 * This is the exact format the Teams chatsvc API expects in the message body's
 * `properties.files` field — a JSON-encoded string (not an array).
 */
export function buildFilesProperty(driveItem: DriveItem): string {
  const baseUrl = extractBaseUrl(driveItem) ?? '';
  const fileName = driveItem.name;
  const fileType = getFileExtension(fileName);
  // Teams uses the file's SharePoint GUID, not its opaque Graph driveItem ID
  // (and never the parent folder's GUID).
  const listItemUniqueId = driveItem.sharepointIds?.listItemUniqueId;
  if (!listItemUniqueId || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(listItemUniqueId)) {
    throw new Error('Missing SharePoint file GUID for Teams attachment');
  }
  const itemId = listItemUniqueId;

  const objectUrl = driveItem.webUrl
    ?? `${baseUrl}Documents/${TEAMS_CHAT_FILES_FOLDER}/${encodeURIComponent(fileName)}`;

  const file = {
    itemid: itemId,
    fileName,
    fileType,
    fileInfo: {
      itemId: null,
      fileUrl: objectUrl,
      siteUrl: baseUrl,
      serverRelativeUrl: '',
      shareUrl: null,
      shareId: null,
    },
    fileChicletState: {
      serviceName: 'p2p',
      state: 'active',
    },
    '@type': 'http://schema.skype.com/File',
    version: 2,
    id: itemId,
    baseUrl,
    objectUrl,
    type: fileType,
    title: fileName,
    state: 'active',
    chicletBreadcrumbs: null,
    providerData: '',
    botFileProperties: {},
    isUploadError: null,
    progressComplete: null,
    filePreview: {
      previewUrl: '',
      previewHeight: 0,
      previewWidth: 0,
    },
    sharepointIds: {
      listId: driveItem.sharepointIds?.listId ?? null,
      listItemUniqueId,
      siteId: driveItem.sharepointIds?.siteId ?? null,
      siteUrl: driveItem.sharepointIds?.siteUrl ?? null,
      webId: driveItem.sharepointIds?.webId ?? null,
    },
    publication: null,
    site: null,
  };

  return JSON.stringify([file]);
}

// ─────────────────────────────────────────────────────────────────────────────
// File Upload
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Uploads a local file to the user's OneDrive "Microsoft Teams Chat Files" folder.
 *
 * Uses a Graph upload session with sequential streamed fragments. Supports files
 * of 2 GiB and larger without loading the file into memory.
 *
 * @param filePath - Absolute or relative path to the local file
 * @returns Upload result with item metadata and the `files` property string
 */
export async function uploadFile(filePath: string): Promise<Result<UploadFileResult>> {
  // Validate auth
  const graphToken = getValidGraphToken();
  if (!graphToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'ACTION REQUIRED: No valid Microsoft Graph token. You MUST call teams_login to authenticate before uploading files.',
      { suggestions: ['Call teams_login to authenticate via browser'] }
    ));
  }

  let file: FileHandle | undefined;
  let driveItem: DriveItem;
  try {
    file = await open(filePath, 'r');
    const info = await file.stat();
    if (!info.isFile() || !Number.isSafeInteger(info.size)) {
      return err(createError(ErrorCode.INVALID_INPUT, 'Expected a regular file with a safely representable size'));
    }
    const response = await uploadContent(file, basename(filePath), info.size, graphToken);
    if (!response.ok) return response;
    driveItem = response.value;
  } catch (error) {
    return err(createError(ErrorCode.INVALID_INPUT,
      `Failed to upload file "${filePath}": ${error instanceof Error ? error.message : String(error)}`,
      { retryable: false }));
  } finally {
    await file?.close();
  }

  const metadata = await getAttachmentDriveItem(driveItem.id);
  if (!metadata.ok) return metadata;
  driveItem = metadata.value;
  const baseUrl = extractBaseUrl(driveItem);
  if (!baseUrl) {
    return err(createError(
      ErrorCode.UNKNOWN,
      `File uploaded successfully but could not determine SharePoint base URL from the response. Item ID: ${driveItem.id}`,
      { retryable: false }
    ));
  }

  const filesProperty = buildFilesProperty(driveItem);

  return ok({
    itemId: driveItem.id,
    fileName: driveItem.name,
    fileType: getFileExtension(driveItem.name),
    fileSize: driveItem.size,
    baseUrl,
    objectUrl: driveItem.webUrl ?? `${baseUrl}Documents/${TEAMS_CHAT_FILES_FOLDER}/${encodeURIComponent(driveItem.name)}`,
    webUrl: driveItem.webUrl,
    listItemUniqueId: driveItem.sharepointIds?.listItemUniqueId,
    filesProperty,
  });
}

/**
 * Uploads multiple files and returns their combined `files` property string.
 *
 * The `files` property in the chatsvc message body is a JSON-encoded array.
 * When multiple files are attached, their entries are merged into a single array.
 *
 * @param filePaths - Array of local file paths
 * @returns Combined `files` property string and per-file upload results
 */
export async function uploadFiles(
  filePaths: string[],
  recipientObjectIds: string[] = []
): Promise<Result<{ filesProperty: string; uploads: UploadFileResult[] }>> {
  const uploads: UploadFileResult[] = [];
  const fileEntries: unknown[] = [];

  for (const filePath of filePaths) {
    const result = await uploadFile(filePath);
    if (!result.ok) {
      return result;
    }
    uploads.push(result.value);

    const shared = await grantAttachmentAccess(result.value.itemId, recipientObjectIds);
    if (!shared.ok) return shared;

    // Parse the filesProperty (which is JSON.stringify([singleFile])) and merge
    try {
      const parsed = JSON.parse(result.value.filesProperty) as unknown[];
      fileEntries.push(...parsed);
    } catch {
      return err(createError(
        ErrorCode.UNKNOWN,
        `Failed to parse files property for uploaded file "${result.value.fileName}"`,
        { retryable: false }
      ));
    }
  }

  return ok({
    filesProperty: JSON.stringify(fileEntries),
    uploads,
  });
}
/** Upload completion omits SharePoint IDs; explicitly request the file metadata. */
export async function getAttachmentDriveItem(itemId: string): Promise<Result<DriveItem>> {
  const token = getValidGraphToken();
  if (!token) return err(createError(ErrorCode.AUTH_REQUIRED, 'Login required to read attachment metadata.'));
  const response = await httpRequest<DriveItem>(GRAPH_FILES_API.item(itemId), {
    headers: { Authorization: `Bearer ${token}` }, redirect: 'error',
  });
  if (!response.ok) return response;
  const item = response.value.data;
  if (item.id !== itemId || !item.name || !item.webUrl || !item.sharepointIds?.listItemUniqueId) {
    return err(createError(ErrorCode.API_ERROR, 'Incomplete SharePoint file metadata; attachment was not sent.'));
  }
  try { buildFilesProperty(item); }
  catch { return err(createError(ErrorCode.API_ERROR, 'Invalid SharePoint file GUID; attachment was not sent.')); }
  return ok(item);
}
