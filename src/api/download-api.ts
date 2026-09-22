/** Download shared SharePoint/OneDrive files using the existing Teams session. */
import { open, unlink } from 'node:fs/promises';
import { isAbsolute } from 'node:path';
import { createHash } from 'node:crypto';
import { getSharePointToken } from '../auth/token-refresh-http.js';
import { sharePointDownloadUrl } from '../utils/api-config.js';
import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import { MAX_DOWNLOAD_BYTES } from '../constants.js';

export interface DownloadFileResult {
  fileName: string;
  outputPath: string;
  size: number;
  contentType: string;
  sha256: string;
}

export async function downloadFile(url: string, outputPath: string): Promise<Result<DownloadFileResult>> {
  if (!isAbsolute(outputPath)) return err(createError(ErrorCode.INVALID_INPUT, 'outputPath must be an absolute file path'));
  let endpoint: ReturnType<typeof sharePointDownloadUrl>;
  try {
    endpoint = sharePointDownloadUrl(url);
  } catch (error) {
    return err(createError(ErrorCode.INVALID_INPUT, error instanceof Error ? error.message : 'Invalid file URL'));
  }
  let token = await getSharePointToken(endpoint.origin);
  if (!token.ok) return token;
  const request = (bearer: string) => httpRequest<Buffer>(endpoint.url, {
    headers: { Authorization: `Bearer ${bearer}` },
    // Do not forward credentials or mistake a sign-in redirect for file contents.
    redirect: 'error',
    responseType: 'buffer',
    maxResponseBytes: MAX_DOWNLOAD_BYTES,
  });
  let response = await request(token.value);
  if (!response.ok && response.error.code === ErrorCode.AUTH_EXPIRED) {
    token = await getSharePointToken(endpoint.origin, true);
    if (!token.ok) return token;
    response = await request(token.value);
  }
  if (!response.ok) return response;
  const bytes = response.value.data;
  // Exclusive creation protects existing files, including symlink targets.
  let file;
  try {
    file = await open(outputPath, 'wx', 0o600);
    await file.writeFile(bytes);
  } catch (error) {
    if (file) await unlink(outputPath).catch(() => {});
    return err(createError(ErrorCode.INVALID_INPUT, `Could not save file: ${error instanceof Error ? error.message : String(error)}`));
  } finally {
    await file?.close();
  }
  return ok({
    fileName: endpoint.fileName,
    outputPath,
    size: bytes.length,
    contentType: response.value.headers.get('content-type') ?? 'application/octet-stream',
    sha256: createHash('sha256').update(bytes).digest('hex'),
  });
}
