import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { downloadImage } from './images-api.js';
import { getMessage } from './chatsvc-messaging.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { clearRateLimitState } from '../utils/http.js';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
vi.mock('./chatsvc-messaging.js', () => ({ getMessage: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireMessageAuthWithConfig: vi.fn() }));
const url = 'https://eu-api.asm.skype.com/v1/objects/0-weu-test/views/imgo';
let directory: string;
function message(imageUrl = url) {
  return ok({ id: '123', conversationId: 'chat', content: '', contentType: 'RichText/Html', sender: { mri: 'user' }, timestamp: '',
    images: [{ index: 0, url: imageUrl, downloadable: true }] });
}
beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'teams-images-'));
  vi.mocked(getMessage).mockReset().mockResolvedValue(message());
  vi.mocked(requireMessageAuthWithConfig).mockReset().mockReturnValue(ok({ auth: { skypeToken: 'test-token', authToken: 'unused', userMri: 'user' }, region: 'emea', baseUrl: 'https://teams.cloud.microsoft' }));
  clearRateLimitState();
});
afterEach(async () => {
  vi.restoreAllMocks(); vi.unstubAllGlobals();
  await rm(directory, { recursive: true, force: true });
});
describe('downloadImage', () => {
  it('downloads the selected message image with ASM auth and returns exact bytes and hash', async () => {
    const bytes = Buffer.from([255, 216, 255, 0, 1, 2]);
    const fetch = vi.fn().mockResolvedValue(new Response(bytes, { headers: { 'content-type': 'image/jpeg' } }));
    vi.stubGlobal('fetch', fetch);
    const output = join(directory, 'image');
    const result = await downloadImage('chat', '123', 0, output);
    expect(result).toEqual(ok({ conversationId: 'chat', messageId: '123', imageIndex: 0, outputPath: output, size: bytes.length, contentType: 'image/jpeg', sha256: createHash('sha256').update(bytes).digest('hex') }));
    expect(await readFile(output)).toEqual(bytes);
    expect(getMessage).toHaveBeenCalledWith('chat', '123');
    expect(fetch).toHaveBeenCalledWith(url, expect.objectContaining({ headers: { Authorization: 'skype_token test-token' }, redirect: 'error' }));
  });
  it('does not overwrite existing files', async () => {
    const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
    const output = join(directory, 'image'); await writeFile(output, 'keep');
    expect((await downloadImage('chat', '123', 0, output)).ok).toBe(false);
    expect(await readFile(output, 'utf8')).toBe('keep'); expect(fetch).not.toHaveBeenCalled();
  });
  it('rejects unsupported message URLs without sending credentials', async () => {
    vi.mocked(getMessage).mockResolvedValue(message('https://evil.example/image'));
    const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
    expect((await downloadImage('chat', '123', 0, join(directory, 'image'))).ok).toBe(false);
    expect(fetch).not.toHaveBeenCalled(); expect(requireMessageAuthWithConfig).not.toHaveBeenCalled();
  });
  it('rejects missing images and invalid input', async () => {
    expect((await downloadImage('chat', '123', 1, join(directory, 'image'))).ok).toBe(false);
    expect((await downloadImage('chat', '123', 0, 'relative')).ok).toBe(false);
    expect(requireMessageAuthWithConfig).not.toHaveBeenCalled();
  });
  it('preserves message access errors', async () => {
    const failure = err(createError(ErrorCode.ACCESS_DENIED, 'denied'));
    vi.mocked(getMessage).mockResolvedValue(failure);
    expect(await downloadImage('chat', '123', 0, join(directory, 'image'))).toEqual(failure);
  });
  it.each([401, 403, 302])('does not leave a file on HTTP %s', async status => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('failure', { status })));
    const output = join(directory, 'image');
    expect((await downloadImage('chat', '123', 0, output)).ok).toBe(false);
    await expect(readFile(output)).rejects.toThrow();
  });
  it('rejects successful HTML login pages', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('<html>login</html>', { headers: { 'content-type': 'text/html' } })));
    const output = join(directory, 'image');
    expect((await downloadImage('chat', '123', 0, output)).ok).toBe(false);
    await expect(readFile(output)).rejects.toThrow();
  });
  it('removes interrupted downloads without retrying partial writes', async () => {
    let chunks = 0;
    const fetch = vi.fn().mockImplementation(() => new Response(new ReadableStream({ pull(controller) {
      if (chunks++ < 3) controller.enqueue(Buffer.from('chunk')); else controller.error(new Error('interrupted'));
    } }), { headers: { 'content-type': 'image/png' } }));
    vi.stubGlobal('fetch', fetch);
    const output = join(directory, 'image');
    expect((await downloadImage('chat', '123', 0, output)).ok).toBe(false);
    expect(fetch).toHaveBeenCalledTimes(1); await expect(readFile(output)).rejects.toThrow();
  });
});
