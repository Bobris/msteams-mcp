import { describe, expect, it } from 'vitest';
import { extractMessageAuth } from './token-extractor.js';
import type { SessionState } from './session-store.js';

const jwt = `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ skypeid: 'orgid:test-user' })).toString('base64url')}.test`;
function state(domain: string): SessionState {
  const cookie = { path: '/', expires: -1, httpOnly: true, secure: true, sameSite: 'None' as const };
  return { origins: [], cookies: [
    { ...cookie, domain: '.asyncgw.teams.microsoft.com', name: 'skypetoken_asm', value: jwt },
    { ...cookie, domain, name: 'authtoken', value: 'Bearer%3Dtest-auth' },
  ] };
}
describe('Teams messaging cookie migration', () => {
  it.each(['teams.microsoft.com', 'teams.cloud.microsoft'])('reads auth cookies split across %s and the messaging gateway', domain => {
    expect(extractMessageAuth(state(domain))).toEqual({ skypeToken: jwt, authToken: 'test-auth', userMri: '8:orgid:test-user' });
  });
  it.each(['teams.cloud.microsoft.evil.test', 'evilteams.microsoft.com'])('rejects unrelated cookie domain %s', domain => {
    expect(extractMessageAuth(state(domain))).toBeNull();
  });
});
