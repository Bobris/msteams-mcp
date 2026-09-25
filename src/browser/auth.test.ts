import { describe, expect, it, vi } from 'vitest';
import type { Page } from 'playwright';
import { getAuthStatus } from './auth.js';

function pageAt(url: string, hasContent = true): Page {
  return {
    url: () => url,
    locator: vi.fn(() => ({ count: async () => hasContent ? 1 : 0 })),
  } as unknown as Page;
}

describe('getAuthStatus', () => {
  it.each([
    'https://teams.cloud.microsoft/v2/',
    'https://teams.microsoft.com/v2/',
    'https://teams.microsoft.us/',
    'https://dod.teams.microsoft.us/',
  ])('recognizes authenticated Teams content at %s', async url => {
    expect((await getAuthStatus(pageAt(url))).isAuthenticated).toBe(true);
  });

  it('waits for the app to load on the new Teams domain', async () => {
    expect((await getAuthStatus(pageAt('https://teams.cloud.microsoft/v2/', false))).isAuthenticated).toBe(false);
  });

  it.each([
    'https://teams.microsoft.com.example.org/',
    'https://example.org/teams.microsoft.com',
    'https://example.org/?next=https://teams.cloud.microsoft/',
    'http://teams.microsoft.com/',
    'not a URL',
  ])('does not recognize an unrelated URL as Teams: %s', async url => {
    expect((await getAuthStatus(pageAt(url))).isAuthenticated).toBe(false);
  });

  it('does not accept content on the Microsoft login page', async () => {
    expect(await getAuthStatus(pageAt('https://login.microsoftonline.com/tenant/oauth2/authorize')))
      .toMatchObject({ isAuthenticated: false, isOnLoginPage: true });
  });
});
