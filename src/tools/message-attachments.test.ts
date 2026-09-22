import { beforeEach, describe, expect, it, vi } from 'vitest';
import { sendMessageTool, SendMessageInputSchema } from './message-tools.js';
import { sendMessage } from '../api/chatsvc-api.js';
import { uploadFiles } from '../api/sharepoint-api.js';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import type { ToolContext } from './index.js';

vi.mock('../api/chatsvc-api.js', async importOriginal => ({
  ...await importOriginal<typeof import('../api/chatsvc-api.js')>(), sendMessage: vi.fn(),
}));
vi.mock('../api/sharepoint-api.js', async importOriginal => ({
  ...await importOriginal<typeof import('../api/sharepoint-api.js')>(), uploadFiles: vi.fn(),
}));
const ctx = {} as ToolContext;
beforeEach(() => vi.clearAllMocks());

describe('message attachments', () => {
  it('passes uploaded file metadata to the outgoing message', async () => {
    const filesProperty = '[{"itemid":"uploaded-id"}]';
    vi.mocked(uploadFiles).mockResolvedValue(ok({ filesProperty, uploads: [] }));
    vi.mocked(sendMessage).mockResolvedValue(ok({ messageId: 'id', conversationId: 'chat', content: 'test', timestamp: 123 }));
    const result = await sendMessageTool.handler(SendMessageInputSchema.parse({
      content: 'test', conversationId: 'chat', attachments: [{ filePath: '/tmp/file.bin' }],
    }), ctx);
    expect(result.success).toBe(true);
    expect(uploadFiles).toHaveBeenCalledWith(['/tmp/file.bin']);
    expect(sendMessage).toHaveBeenCalledWith('chat', 'test', expect.objectContaining({ files: filesProperty }));
  });
  it('does not send a message after an upload failure', async () => {
    vi.mocked(uploadFiles).mockResolvedValue(err(createError(ErrorCode.NETWORK_ERROR, 'Upload failed')));
    const result = await sendMessageTool.handler(SendMessageInputSchema.parse({
      content: 'test', attachments: [{ filePath: '/tmp/file.bin' }],
    }), ctx);
    expect(result.success).toBe(false);
    expect(sendMessage).not.toHaveBeenCalled();
  });
  it('rejects scheduled attachments before starting uploads', async () => {
    const result = await sendMessageTool.handler(SendMessageInputSchema.parse({
      content: 'test', scheduleAt: '2030-01-01T00:00:00Z', attachments: [{ filePath: '/tmp/file.bin' }],
    }), ctx);
    expect(result.success).toBe(false);
    expect(uploadFiles).not.toHaveBeenCalled();
    expect(sendMessage).not.toHaveBeenCalled();
  });
});
