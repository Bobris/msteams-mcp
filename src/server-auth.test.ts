import { it, expect, vi } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
vi.mock('./tools/registry.js', () => ({ invokeTool: vi.fn(), getToolDefinitions: () => [] }));
vi.mock('./auth/token-refresh.js', () => ({ refreshTokensViaBrowser: vi.fn() }));
import { invokeTool } from './tools/registry.js';
import { refreshTokensViaBrowser } from './auth/token-refresh.js';
import { TeamsServer } from './server.js';
import { createError, ErrorCode } from './types/errors.js';

it('falls back to browser when the original operation still rejects a partial HTTP refresh', async () => {
  const failure = { success: false as const, error: createError(ErrorCode.AUTH_EXPIRED, 'Skype expired') };
  vi.mocked(invokeTool).mockResolvedValueOnce(failure).mockResolvedValueOnce(failure)
    .mockResolvedValueOnce({ success: true, data: { recovered: true } });
  vi.mocked(refreshTokensViaBrowser).mockResolvedValue({ ok: true, value: {
    newExpiry: new Date(), previousExpiry: new Date(), minutesGained: 60, refreshNeeded: true, method: 'http',
  } });
  const server = await new TeamsServer().createServer();
  const client = new Client({ name: 'test', version: '1' });
  const [a, b] = InMemoryTransport.createLinkedPair();
  try {
    await server.connect(a);
    await client.connect(b);
    const result = await client.callTool({ name: 'teams_get_thread', arguments: {} });
    expect(result.isError).not.toBe(true);
    expect(refreshTokensViaBrowser).toHaveBeenNthCalledWith(2, true);
    expect(invokeTool).toHaveBeenCalledTimes(3);
  } finally {
    await client.close();
    await server.close();
  }
});
