/** Authenticated, read-only downloads of inline Teams ASM images. */
import { open, unlink } from 'node:fs/promises';
import { isAbsolute } from 'node:path';
import { createHash } from 'node:crypto';
import { getMessage } from './chatsvc-messaging.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { isTeamsImageUrl } from '../utils/parsers-images.js';
import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';

export interface DownloadImageResult {
  conversationId: string;
  messageId: string;
  imageIndex: number;
  outputPath: string;
  size: number;
  contentType: string;
  sha256: string;
}

export async function downloadImage(conversationId: string, messageId: string, imageIndex: number, outputPath: string): Promise<Result<DownloadImageResult>> {
  if (!isAbsolute(outputPath) || !Number.isInteger(imageIndex) || imageIndex < 0) {
    return err(createError(ErrorCode.INVALID_INPUT, 'Use an absolute outputPath and a non-negative integer imageIndex'));
  }
  const message = await getMessage(conversationId, messageId);
  if (!message.ok) return message;
  const image = message.value.images?.[imageIndex];
  if (!image) return err(createError(ErrorCode.INVALID_INPUT, 'Image index not found; read the message images array first'));
  if (!isTeamsImageUrl(image.url)) return err(createError(ErrorCode.INVALID_INPUT, 'This image is not on a supported Teams ASM endpoint'));
  const auth = requireMessageAuthWithConfig();
  if (!auth.ok) return auth;
  let file;
  let complete = false;
  try {
    file = await open(outputPath, 'wx', 0o600);
    const destination = file;
    const response = await httpRequest(image.url, {
      headers: { Authorization: `skype_token ${auth.value.auth.skypeToken}` },
      redirect: 'error',
      maxRetries: 1,
      consumeResponse: async (response, resetTimeout) => {
        const contentType = response.headers.get('content-type')?.split(';')[0].trim().toLowerCase();
        if (!contentType?.startsWith('image/')) {
          await response.body?.cancel();
          throw new Error('Teams returned non-image content');
        }
        const reader = response.body?.getReader();
        if (!reader) throw new Error('Teams returned an empty image response');
        const hash = createHash('sha256');
        let size = 0;
        try {
          while (true) {
            const next = await reader.read();
            if (next.done) break;
            resetTimeout();
            await destination.writeFile(next.value);
            hash.update(next.value);
            size += next.value.byteLength;
            resetTimeout();
          }
        } finally {
          await reader.cancel().catch(() => {});
          reader.releaseLock();
        }
        if (!size) throw new Error('Teams returned an empty image');
        return { size, contentType, sha256: hash.digest('hex') };
      },
    });
    if (!response.ok) return response;
    await file.close();
    complete = true;
    return ok({ conversationId, messageId, imageIndex, outputPath, ...response.value.data });
  } catch (error) {
    return err(createError(ErrorCode.INVALID_INPUT, `Could not save image: ${error instanceof Error ? error.message : String(error)}`));
  } finally {
    if (file && !complete) {
      await file.close().catch(() => {});
      await unlink(outputPath).catch(() => {});
    }
  }
}
