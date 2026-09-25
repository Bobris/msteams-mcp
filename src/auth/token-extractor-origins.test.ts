import { describe, it, expect } from 'vitest';
import { extractSubstrateToken } from './token-extractor.js';
import type { SessionState } from './session-store.js';

function entry(account: string, seconds: number) {
  const secret = `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ exp: Math.floor(Date.now() / 1000) + seconds })).toString('base64url')}.test`;
  return { name: 'token', value: JSON.stringify({ credentialType: 'AccessToken', homeAccountId: account, target: 'https://substrate.office.com/SubstrateSearch', secret }) };
}
describe('token extraction across Teams origins', () => {
  it('reads the renewed token on the old host when the new host still has an expired token', () => {
    const fresh = entry('user.tenant', 3600);
    const state: SessionState = { cookies: [], origins: [
      { origin: 'https://teams.cloud.microsoft', localStorage: [entry('user.tenant', -60)] },
      { origin: 'https://teams.microsoft.com', localStorage: [fresh] },
    ] };
    expect(extractSubstrateToken(state)?.token).toBe(JSON.parse(fresh.value).secret);
  });
  it.each(['different-account', 'untrusted-host'])('does not select a token from %s', scenario => {
    const state: SessionState = { cookies: [], origins: [
      { origin: 'https://teams.cloud.microsoft', localStorage: [entry('user.tenant', -60)] },
      { origin: scenario === 'untrusted-host' ? 'https://evil.test' : 'https://teams.microsoft.com', localStorage: [entry(scenario === 'different-account' ? 'other.tenant' : 'user.tenant', 3600)] },
    ] };
    expect(extractSubstrateToken(state)).toBeNull();
  });
});
