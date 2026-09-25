/** Grant chat recipients access before publishing OneDrive attachment references. */
import { getValidGraphToken } from '../auth/token-extractor.js';
import { requireMessageAuth } from '../utils/auth-guards.js';
import { httpRequest, type HttpResponse } from '../utils/http.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import { GRAPH_FILES_API } from '../utils/api-config.js';

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const failure = (message: string) => err(createError(ErrorCode.API_ERROR, message, { retryable: false }));

/** Self notes stay private; resolve every recipient before starting any upload. */
export async function getAttachmentRecipients(conversationId: string): Promise<Result<string[]>> {
  if (conversationId === '48:notes') return ok([]);
  const auth = requireMessageAuth();
  if (!auth.ok) return auth;
  const self = auth.value.userMri.replace(/^8:orgid:/, '').toLowerCase();
  if (!GUID.test(self)) return failure('Cannot identify the attachment owner.');

  // A new 1:1 chat may not exist yet. Its deterministic ID contains both members.
  const direct = /^19:([^_]+)_([^@]+)@unq\.gbl\.spaces$/.exec(conversationId);
  if (direct) {
    const members = direct.slice(1).map(id => id.toLowerCase());
    if (!members.every(id => GUID.test(id)) || !members.includes(self)) {
      return failure('The 1:1 attachment destination does not identify the current user.');
    }
    return ok([...new Set(members.filter(id => id !== self))]);
  }
  if (!/^19:.+@thread\.v2$/.test(conversationId)) {
    return failure('File attachments currently support private chats and self notes, not channels.');
  }
  const token = getValidGraphToken();
  if (!token) return err(createError(ErrorCode.AUTH_REQUIRED, 'Login required to resolve attachment recipients.'));
  const recipients = new Set<string>();
  const seen = new Set<string>();
  let url: string | undefined = GRAPH_FILES_API.chatMembers(conversationId);
  while (url) {
    if (seen.has(url) || !url.startsWith(GRAPH_FILES_API.base + '/chats/')) {
      return failure('Invalid chat membership pagination response.');
    }
    seen.add(url);
    const response: Result<HttpResponse<{ value?: { userId?: string }[]; '@odata.nextLink'?: string }>> = await httpRequest(url, {
      headers: { Authorization: `Bearer ${token}` }, redirect: 'error',
    });
    if (!response.ok) return response;
    if (!response.value.data.value?.length) return failure('No chat members returned; attachments were not uploaded.');
    for (const member of response.value.data.value) {
      if (!member.userId || !GUID.test(member.userId)) return failure('Cannot resolve every chat member for file sharing.');
      recipients.add(member.userId.toLowerCase());
    }
    url = response.value.data['@odata.nextLink'];
  }
  if (!recipients.delete(self) || !recipients.size) return failure('Incomplete chat membership; attachments were not uploaded.');
  return ok([...recipients]);
}

interface Permission {
  roles?: string[];
  error?: unknown;
  grantedToV2?: { user?: { id?: string } };
  grantedTo?: { user?: { id?: string } };
}

/** Direct, signed-in read access only. Never create an anonymous/organization link. */
export async function grantAttachmentAccess(itemId: string, recipients: string[]): Promise<Result<void>> {
  if (!recipients.length) return ok(undefined);
  if (!recipients.every(id => GUID.test(id))) return failure('Invalid attachment recipient ID.');
  const token = getValidGraphToken();
  if (!token) return err(createError(ErrorCode.AUTH_REQUIRED, 'Login required to share the uploaded attachment.'));
  const unique = [...new Set(recipients.map(id => id.toLowerCase()))];
  for (let start = 0; start < unique.length; start += 20) {
    const batch = unique.slice(start, start + 20);
    const response = await httpRequest<{ value?: Permission[] }>(GRAPH_FILES_API.invite(itemId), {
      method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ recipients: batch.map(objectId => ({ objectId })), roles: ['read'], requireSignIn: true, sendInvitation: false }),
      maxRetries: 1, redirect: 'error',
    });
    if (!response.ok) return response;
    const granted = new Set((response.value.data.value ?? [])
      .filter(p => !p.error && p.roles?.some(role => ['read', 'write', 'owner'].includes(role)))
      .map(p => (p.grantedToV2?.user?.id ?? p.grantedTo?.user?.id ?? '').toLowerCase()));
    if (response.value.status === 207 || !batch.every(id => granted.has(id))) {
      return failure(`Could not confirm access for every recipient of uploaded file ${itemId}; message was not sent.`);
    }
  }
  return ok(undefined);
}
