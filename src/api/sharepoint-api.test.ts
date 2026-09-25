import { beforeEach, afterEach, describe, expect, it, vi } from 'vitest';
import { mkdtemp, open, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { uploadFile, uploadFiles, buildFilesProperty } from './sharepoint-api.js';
import { getValidGraphToken } from '../auth/token-extractor.js';
import { clearRateLimitState } from '../utils/http.js';
import { UPLOAD_CHUNK_BYTES, UPLOAD_READ_BYTES } from '../constants.js';

vi.mock('../auth/token-extractor.js', () => ({ getValidGraphToken: vi.fn() }));
let directory: string;
const sessionUrl = 'https://example.sharepoint.com/upload-session';
const json = (data: unknown, status = 200) => new Response(JSON.stringify(data), {
  status, headers: { 'Content-Type': 'application/json' },
});
const item = (size: number) => ({
  id: 'item-id', name: 'file.bin', size,
  sharepointIds: { listItemUniqueId: 'a49bccb4-6ce5-4af0-85c8-3cb57fbf100a' },
  webUrl: 'https://example-my.sharepoint.com/personal/user/Documents/Microsoft%20Teams%20Chat%20Files/file.bin',
});
beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'teams-upload-'));
  vi.mocked(getValidGraphToken).mockReturnValue('graph-test-token');
  clearRateLimitState();
});
afterEach(async () => { vi.unstubAllGlobals(); await rm(directory, { recursive: true, force: true }); });

/** Drain each actual request stream, checking all range and authentication boundaries. */
function mockUpload(size: number, inspect?: (bytes: Uint8Array, offset: number) => void) {
  let received = 0;
  let parts = 0;
  let largestRead = 0;
  const fetchMock = vi.fn().mockImplementation(async (url: string, options: RequestInit) => {
    const headers = new Headers(options.headers);
    if (options.method === 'GET' || !options.method) {
      expect(url).toContain('/me/drive/items/item-id?$select=');
      expect(headers.get('Authorization')).toBe('Bearer graph-test-token');
      return json(item(size));
    }
    if (options.method === 'POST') {
      expect(url).toContain(':/createUploadSession');
      expect(headers.get('Authorization')).toBe('Bearer graph-test-token');
      expect(JSON.parse(options.body as string).item['@microsoft.graph.conflictBehavior']).toBe('rename');
      return json({ uploadUrl: sessionUrl });
    }
    expect(url).toBe(sessionUrl);
    expect(headers.has('Authorization')).toBe(false);
    expect(options.method).toBe('PUT');
    expect(options.body).toBeInstanceOf(ReadableStream);
    const end = Math.min(received + UPLOAD_CHUNK_BYTES, size) - 1;
    expect(headers.get('Content-Range')).toBe(`bytes ${received}-${end}/${size}`);
    expect(headers.get('Content-Length')).toBe(String(end - received + 1));
    const reader = (options.body as ReadableStream<Uint8Array>).getReader();
    try {
      while (true) {
        const next = await reader.read();
        if (next.done) break;
        largestRead = Math.max(largestRead, next.value.byteLength);
        inspect?.(next.value, received);
        received += next.value.byteLength;
      }
    } finally { reader.releaseLock(); }
    expect(received).toBe(end + 1);
    parts++;
    return received === size ? json(item(size), 201) : json({ nextExpectedRanges: [`${received}-${parts % 2 ? size - 1 : ''}`] }, 202);
  });
  vi.stubGlobal('fetch', fetchMock);
  return { fetchMock, stats: () => ({ received, parts, largestRead }) };
}

describe('streaming Graph uploads', () => {
  it('streams a file larger than 2 GiB using bounded reads and correct 64-bit ranges', async () => {
    const size = 2 * 1024 * 1024 * 1024 + 123;
    const path = join(directory, 'large.bin');
    const file = await open(path, 'wx');
    await file.truncate(size); // Sparse fixture: reads real bytes without allocating 2 GiB of RAM/disk.
    const markers = [0, UPLOAD_CHUNK_BYTES - 1, UPLOAD_CHUNK_BYTES, 2 ** 31, size - 1];
    for (const position of markers) await file.write(Buffer.from([0xa5]), 0, 1, position);
    await file.close();
    const seen = new Set<number>();
    const mock = mockUpload(size, (bytes, start) => {
      for (const position of markers) {
        if (position >= start && position < start + bytes.length) {
          expect(bytes[position - start]).toBe(0xa5);
          seen.add(position);
        }
      }
    });
    const result = await uploadFile(path);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.fileSize).toBe(size);
    expect(mock.stats()).toEqual({ received: size, parts: Math.ceil(size / UPLOAD_CHUNK_BYTES), largestRead: UPLOAD_READ_BYTES });
    expect(seen.size).toBe(markers.length);
  }, 60000);

  it('preserves binary content and the final short fragment', async () => {
    const bytes = Buffer.alloc(UPLOAD_CHUNK_BYTES + 37);
    for (let i = 0; i < bytes.length; i++) bytes[i] = i % 251;
    const path = join(directory, 'binary.bin');
    await writeFile(path, bytes);
    const mock = mockUpload(bytes.length, (part, offset) => {
      expect(Buffer.from(part).equals(bytes.subarray(offset, offset + part.length))).toBe(true);
    });
    expect((await uploadFile(path)).ok).toBe(true);
    expect(mock.stats().parts).toBe(2);
  });

  it('uploads empty files without an invalid session byte range', async () => {
    const path = join(directory, 'empty.bin');
    await writeFile(path, '');
    const fetchMock = vi.fn().mockImplementation(async () => json(item(0), 201));
    vi.stubGlobal('fetch', fetchMock);
    expect((await uploadFile(path)).ok).toBe(true);
    expect(fetchMock).toHaveBeenCalledWith(expect.stringContaining(':/content'), expect.objectContaining({ method: 'PUT', body: new Uint8Array(0) }));
  });

  it('cancels an incomplete session after an upload error without replaying a stream', async () => {
    const path = join(directory, 'file');
    await writeFile(path, 'content');
    const fetchMock = vi.fn().mockResolvedValueOnce(json({ uploadUrl: sessionUrl }))
      .mockResolvedValueOnce(json({ error: 'failed' }, 500)).mockResolvedValueOnce(new Response(null, { status: 204 }));
    vi.stubGlobal('fetch', fetchMock);
    expect((await uploadFile(path)).ok).toBe(false);
    expect(fetchMock).toHaveBeenCalledTimes(3);
    expect(fetchMock).toHaveBeenLastCalledWith(sessionUrl, expect.objectContaining({ method: 'DELETE' }));
  });

  it('cancels the session when the server returns an unexpected next range', async () => {
    const path = join(directory, 'file');
    await writeFile(path, Buffer.alloc(UPLOAD_CHUNK_BYTES + 1));
    const fetchMock = vi.fn().mockResolvedValueOnce(json({ uploadUrl: sessionUrl }))
      .mockResolvedValueOnce(json({ nextExpectedRanges: ['0-'] }, 202)).mockResolvedValueOnce(new Response(null, { status: 204 }));
    vi.stubGlobal('fetch', fetchMock);
    const result = await uploadFile(path);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('next byte range');
    expect(fetchMock).toHaveBeenLastCalledWith(sessionUrl, expect.objectContaining({ method: 'DELETE' }));
  });

  it('rejects missing files and directories before creating an upload session', async () => {
    vi.stubGlobal('fetch', vi.fn());
    expect((await uploadFile(join(directory, 'missing'))).ok).toBe(false);
    expect((await uploadFile(directory)).ok).toBe(false);
    expect(fetch).not.toHaveBeenCalled();
  });

  it('stops multiple uploads at the first failure', async () => {
    const path = join(directory, 'first');
    await writeFile(path, '');
    vi.stubGlobal('fetch', vi.fn().mockImplementation(async () => json(item(0), 201)));
    const result = await uploadFiles([path, join(directory, 'missing'), path]);
    expect(result.ok).toBe(false);
    expect(fetch).toHaveBeenCalledTimes(2);
  });
});


describe('Teams attachment identity', () => {
  it('uses the file SharePoint GUID for every Teams identity, never the Graph ID', () => {
    const entry = JSON.parse(buildFilesProperty(item(42)))[0];
    expect(entry.itemid).toBe(item(42).sharepointIds.listItemUniqueId);
    expect(entry.id).toBe(entry.itemid);
    expect(entry.sharepointIds.listItemUniqueId).toBe(entry.itemid);
    expect(entry.permissionScope).not.toBe('anonymous');
  });
  it('rejects a missing file GUID even when the parent has a valid GUID', () => {
    expect(() => buildFilesProperty({ ...item(42), sharepointIds: undefined,
      parentReference: { sharepointIds: item(42).sharepointIds } })).toThrow('SharePoint file GUID');
  });
  it('hydrates upload metadata before producing a Teams attachment', async () => {
    const path = join(directory, 'empty'); await writeFile(path, '');
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(json({ ...item(0), sharepointIds: undefined }, 201))
      .mockResolvedValueOnce(json(item(0))));
    const result = await uploadFile(path);
    expect(result.ok).toBe(true);
    if (result.ok) expect(JSON.parse(result.value.filesProperty)[0].itemid).toBe(item(0).sharepointIds.listItemUniqueId);
  });
  it('returns failure when Graph cannot provide file identity', async () => {
    const path = join(directory, 'empty'); await writeFile(path, '');
    vi.stubGlobal('fetch', vi.fn().mockImplementation(async () => json({ ...item(0), sharepointIds: undefined })));
    expect((await uploadFile(path)).ok).toBe(false);
  });
});
