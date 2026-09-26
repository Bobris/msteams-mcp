import { beforeEach, expect, it, vi } from 'vitest';
import { refreshTokensViaBrowser } from './token-refresh.js';
import { requireGraphTokenAsync } from '../utils/auth-guards.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { extractSubstrateToken, getValidGraphToken } from './token-extractor.js';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';

vi.mock('./token-refresh-http.js', () => ({ refreshTokensViaHttp: vi.fn() }));
vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(), getValidGraphToken: vi.fn(),
}));
vi.mock('../browser/context.js', () => ({
  createBrowserContext: vi.fn(() => { throw new Error('Unexpected browser login'); }),
}));

const refreshed = ok({ tokensRefreshed: 1, skypeTokenRefreshed: false, refreshTokenRotated: false });

beforeEach(() => { vi.resetAllMocks(); });

it('uses a fresh Graph token without any refresh', async () => {
  vi.mocked(getValidGraphToken).mockReturnValue('cached');
  expect(await requireGraphTokenAsync()).toEqual(ok('cached'));
  expect(refreshTokensViaHttp).not.toHaveBeenCalled();
});

it('deduplicates concurrent Graph refresh and isolates auth failures from Teams auto-login', async () => {
  vi.mocked(refreshTokensViaHttp).mockResolvedValue(err(createError(ErrorCode.AUTH_EXPIRED, 'expired')));
  const results = await Promise.all([requireGraphTokenAsync(), requireGraphTokenAsync()]);
  expect(refreshTokensViaHttp).toHaveBeenCalledExactlyOnceWith('graph');
  for (const result of results) {
    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_INTERACTION_REQUIRED, retryable: false } });
  }
});

it('serializes Graph behind a core refresh and recovers without an unexpired access token', async () => {
  let finish!: (value: typeof refreshed) => void;
  vi.mocked(extractSubstrateToken).mockReturnValueOnce(null).mockReturnValue({
    token: 'substrate', expiry: new Date(Date.now() + 3600_000),
  });
  vi.mocked(refreshTokensViaHttp)
    .mockImplementationOnce(() => new Promise(resolve => { finish = resolve; }))
    .mockImplementationOnce(async () => {
      vi.mocked(getValidGraphToken).mockReturnValue('graph');
      return refreshed;
    });
  const core = refreshTokensViaBrowser();
  const sameCore = refreshTokensViaBrowser();
  const graph = requireGraphTokenAsync();
  await vi.waitFor(() => { expect(refreshTokensViaHttp).toHaveBeenCalledTimes(1); });
  finish(refreshed);
  expect(await core).toMatchObject({ ok: true });
  expect(await sameCore).toEqual(await core);
  expect(await graph).toEqual(ok('graph'));
  expect(vi.mocked(refreshTokensViaHttp).mock.calls).toEqual([[], ['graph']]);
});
