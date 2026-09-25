import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { getAttachmentRecipients, grantAttachmentAccess } from './attachment-sharing.js';
import { requireMessageAuth } from '../utils/auth-guards.js';
import { getValidGraphToken } from '../auth/token-extractor.js';
import { ok } from '../types/result.js';
import { clearRateLimitState } from '../utils/http.js';
vi.mock('../utils/auth-guards.js', () => ({ requireMessageAuth: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({ getValidGraphToken: vi.fn() }));
const self = '4d1607dd-0007-4a3e-9dad-c7a5d767e418';
const other = '7fde7ace-3411-4cf2-9723-ef5416a4d5f1';
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), {status, headers: {'Content-Type': 'application/json'}});
beforeEach(() => {
  vi.clearAllMocks(); clearRateLimitState();
  vi.mocked(requireMessageAuth).mockReturnValue(ok({userMri: `8:orgid:${self}`, skypeToken:'skype', authToken:'auth'}));
  vi.mocked(getValidGraphToken).mockReturnValue('test-token');
});
afterEach(() => vi.unstubAllGlobals());
describe('attachment recipients and grants', () => {
  it('keeps self notes private and resolves a new direct chat without uploading', async () => {
    vi.stubGlobal('fetch', vi.fn());
    expect(await getAttachmentRecipients('48:notes')).toEqual(ok([]));
    expect(await getAttachmentRecipients(`19:${self}_${other}@unq.gbl.spaces`)).toEqual(ok([other]));
    expect(fetch).not.toHaveBeenCalled();
  });
  it('rejects a direct chat without the sender, and channels', async () => {
    expect((await getAttachmentRecipients(`19:${other}_${other}@unq.gbl.spaces`)).ok).toBe(false);
    expect((await getAttachmentRecipients('19:channel@thread.tacv2')).ok).toBe(false);
  });
  it('reads every page of group members and excludes the sender', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(json({value:[{userId:self}], '@odata.nextLink':'https://graph.microsoft.com/v1.0/chats/group/members?$skiptoken=next'}))
      .mockResolvedValueOnce(json({value:[{userId:other}]})));
    expect(await getAttachmentRecipients('19:group@thread.v2')).toEqual(ok([other]));
  });
  it('does not forward the token to a foreign pagination host', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(json({value:[{userId:self}], '@odata.nextLink':'https://untrusted.example/members'}));
    vi.stubGlobal('fetch', fetchMock);
    expect((await getAttachmentRecipients('19:group@thread.v2')).ok).toBe(false);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it('grants only signed-in read access without sending an email invitation', async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(json({value:[{roles:['read'],grantedToV2:{user:{id:other}}}]}));
    vi.stubGlobal('fetch', fetchMock);
    expect(await grantAttachmentAccess('item-id',[other,other])).toEqual(ok(undefined));
    const [url, options] = fetchMock.mock.calls[0];
    expect(url).toBe('https://graph.microsoft.com/v1.0/me/drive/items/item-id/invite');
    expect(JSON.parse(options.body)).toEqual({recipients:[{objectId:other}],roles:['read'],requireSignIn:true,sendInvitation:false});
  });
  it.each([200,207])('rejects incomplete or partial grant responses (%i)', async status => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValueOnce(json({value:[{roles:['read'],grantedToV2:{user:{id:self}}}]},status)));
    expect((await grantAttachmentAccess('item-id',[other])).ok).toBe(false);
  });
});
