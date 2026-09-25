import { it, expect, vi } from 'vitest';
vi.mock('../auth/session-store.js', () => ({ hasSessionState: vi.fn(), isSessionLikelyExpired: vi.fn(), clearSessionState: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({
  getSubstrateTokenStatus: vi.fn(() => ({ hasToken: true, minutesRemaining: 60 })),
  getMessageAuthStatus: vi.fn(), extractMessageAuth: vi.fn(), extractCsaToken: vi.fn(), clearTokenCache: vi.fn(),
}));
vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn(), forceNewLogin: vi.fn(), getAuthStatus: vi.fn() }));
import { getTool } from './registry.js';
import { clearSessionState } from '../auth/session-store.js';
import { createBrowserContext } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import type { ToolContext } from './index.js';

it('recover bypasses the cached-token fast path without deleting SSO cookies', async () => {
  const clearCookies = vi.fn();
  const manager = { page: {}, context: { clearCookies } };
  vi.mocked(createBrowserContext).mockResolvedValue(manager as unknown as Awaited<ReturnType<typeof createBrowserContext>>);
  const ctx = { server: { getBrowserManager: () => null, setBrowserManager: vi.fn(), resetBrowserState: vi.fn(), markInitialised: vi.fn() } } as unknown as ToolContext;
  const login = getTool('teams_login')!;
  const result = await login.handler(login.schema.parse({ recover: true }), ctx);
  expect(result.success).toBe(true);
  expect(ensureAuthenticated).toHaveBeenCalled();
  expect(clearSessionState).not.toHaveBeenCalled();
  expect(clearCookies).not.toHaveBeenCalled();
});
