import { beforeEach, describe, expect, it, vi } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { TeamsServer } from '../server.js';
import { getMessage, getThreadMessages } from './chatsvc-messaging.js';
import { httpRequest } from '../utils/http.js';
import { ok } from '../types/result.js';
vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireMessageAuthWithConfig: () => ok({ auth: { skypeToken: 'token', userMri: 'user' }, region: 'emea', baseUrl: 'https://teams.cloud.microsoft' }),
  getTenantId: () => 'tenant', getTeamsBaseUrl: () => 'https://teams.cloud.microsoft',
}));
const raw = { id: '1790672312942', messagetype: 'RichText/Html', from: 'other', originalarrivaltime: '2026-09-29T08:58:32Z',
  content: '<p>Screenshot</p><img src="https://eu-api.asm.skype.com/v1/objects/test/views/imgo" alt="image" width="450" height="250">' };
beforeEach(() => vi.mocked(httpRequest).mockReset());
describe('message image metadata', () => {
  it('keeps images in single-message responses', async () => {
    vi.mocked(httpRequest).mockResolvedValue(ok({ status: 200, headers: new Headers(), data: raw }));
    const result = await getMessage('chat', raw.id);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.content).toBe('Screenshot');
      expect(result.value.images).toEqual([{ index: 0, url: 'https://eu-api.asm.skype.com/v1/objects/test/views/imgo', alt: 'image', width: 450, height: 250, downloadable: true }]);
    }
  });
  it('keeps image-only thread messages and omits images for text-only messages', async () => {
    vi.mocked(httpRequest).mockResolvedValue(ok({ status: 200, headers: new Headers(), data: { messages: [
      { ...raw, content: raw.content.replace('<p>Screenshot</p>', '') },
      { ...raw, id: '1790672312943', content: 'text' },
    ] } }));
    const result = await getThreadMessages('chat');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.messages).toHaveLength(2);
      expect(result.value.messages.find(m => m.id === raw.id)?.images).toHaveLength(1);
      expect(result.value.messages.find(m => m.content === 'text')?.images).toBeUndefined();
    }
  });
});

// Exercise the public MCP formatter as well: the API result alone is insufficient.
it('exposes images and the download tool through MCP', async () => {
  vi.mocked(httpRequest).mockResolvedValue(ok({ status: 200, headers: new Headers(), data: raw }));
  const server = await new TeamsServer().createServer();
  const client = new Client({ name: 'image-test', version: '1' });
  const [a, b] = InMemoryTransport.createLinkedPair();
  try {
    await server.connect(a); await client.connect(b);
    const tools = await client.listTools();
    expect(tools.tools.some(tool => tool.name === 'teams_download_image')).toBe(true);
    const result = await client.callTool({ name: 'teams_get_message', arguments: { conversationId: 'chat', messageId: raw.id } });
    const content = result.content as Array<{ type: string; text?: string }>;
    expect(JSON.parse(content[0].text!).images[0].downloadable).toBe(true);
    const invalid = await client.callTool({ name: 'teams_download_image', arguments: {
      conversationId: 'chat', messageId: raw.id, outputPath: '/tmp/unused-image', imageIndex: 0.5,
    } });
    expect(invalid.isError).toBe(true);
  } finally { await client.close(); await server.close(); }
});
