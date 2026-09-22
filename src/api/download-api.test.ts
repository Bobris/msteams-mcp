import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { downloadFile } from './download-api.js';
import { getSharePointToken } from '../auth/token-refresh-http.js';
import { clearRateLimitState } from '../utils/http.js';
import { ok } from '../types/result.js';
import { MAX_DOWNLOAD_BYTES } from '../constants.js';

vi.mock('../auth/token-refresh-http.js', () => ({ getSharePointToken: vi.fn() }));
const url = 'https://example-my.sharepoint.com/personal/test/Documents/file.md';
let directory: string;
beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'teams-download-'));
  vi.mocked(getSharePointToken).mockReset().mockResolvedValue(ok('test-token'));
  clearRateLimitState();
});
afterEach(async () => {
  vi.unstubAllGlobals();
  await rm(directory, { recursive: true, force: true });
});

describe('downloadFile', () => {
  it('preserves binary bytes, even with a JSON content type', async () => {
    const bytes = Buffer.from([0, 255, 254, 128, 10]);
    const fetchMock = vi.fn().mockResolvedValue(new Response(bytes, { headers: { 'content-type': 'application/json' } }));
    vi.stubGlobal('fetch', fetchMock);
    const output = join(directory, 'file');
    const result = await downloadFile(url, output);
    expect(result.ok).toBe(true);
    expect(await readFile(output)).toEqual(bytes);
    expect(fetchMock).toHaveBeenCalledWith(expect.stringContaining('/_api/web/GetFileByServerRelativePath('), expect.objectContaining({
      redirect: 'error', headers: { Authorization: 'Bearer test-token' },
    }));
  });
  it('does not overwrite an existing file', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('new')));
    const output = join(directory, 'file');
    await writeFile(output, 'original');
    expect((await downloadFile(url, output)).ok).toBe(false);
    expect(await readFile(output, 'utf8')).toBe('original');
  });
  it('rejects unsupported URLs before accessing credentials', async () => {
    expect((await downloadFile('https://example.org/file.md', join(directory, 'file'))).ok).toBe(false);
    expect(getSharePointToken).not.toHaveBeenCalled();
  });
  it('rejects relative output paths', async () => {
    expect((await downloadFile(url, 'file.md')).ok).toBe(false);
    expect(getSharePointToken).not.toHaveBeenCalled();
  });
  it('refreshes the host token once after a 401', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(new Response('expired', { status: 401 }))
      .mockResolvedValueOnce(new Response('# markdown'));
    vi.stubGlobal('fetch', fetchMock);
    expect((await downloadFile(url, join(directory, 'file'))).ok).toBe(true);
    expect(getSharePointToken).toHaveBeenLastCalledWith('https://example-my.sharepoint.com', true);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
  it('does not create a file when the server denies access', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('forbidden', { status: 403 })));
    const output = join(directory, 'file');
    expect((await downloadFile(url, output)).ok).toBe(false);
    await expect(readFile(output)).rejects.toThrow();
  });
  it('enforces the byte limit even without Content-Length', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(new ReadableStream({
      start(controller) { controller.enqueue(new Uint8Array(MAX_DOWNLOAD_BYTES + 1)); controller.close(); },
    }))));
    const output = join(directory, 'file');
    const result = await downloadFile(url, output);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('download limit');
    await expect(readFile(output)).rejects.toThrow();
  });
});
