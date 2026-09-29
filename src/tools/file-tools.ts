/**
 * File-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { downloadImage } from '../api/images-api.js';
import { downloadFile } from '../api/graph-files-api.js';
import { getSharedFiles } from '../api/files-api.js';
import { uploadFile } from '../api/sharepoint-api.js';
import {
  DEFAULT_FILES_PAGE_SIZE,
  MAX_FILES_PAGE_SIZE,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const GetSharedFilesInputSchema = z.object({
  conversationId: z.string().min(1),
  pageSize: z.number().min(1).max(MAX_FILES_PAGE_SIZE).optional().default(DEFAULT_FILES_PAGE_SIZE),
  skipToken: z.string().optional(),
});

export const UploadFileInputSchema = z.object({
  filePath: z.string().min(1, 'File path cannot be empty'),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const getSharedFilesToolDefinition: Tool = {
  name: 'teams_get_shared_files',
  description: 'Get files and links shared in a Teams conversation. Inline images are separate: use teams_get_thread or teams_get_message and teams_download_image. Returns file names, URLs, extensions, sizes, and who shared them. Works for channels, group chats, 1:1 chats, and meeting chats. Use the conversationId from other tools (teams_get_favorites, teams_search, teams_find_channel, teams_get_chat). Supports pagination via skipToken for conversations with many files. Pass a File item webUrl to teams_download_file to download its contents.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to get shared files for (e.g., "19:abc@thread.tacv2" for a channel, or a chat conversation ID).',
      },
      pageSize: {
        type: 'number',
        description: `Number of items per page (default: ${DEFAULT_FILES_PAGE_SIZE}, max: ${MAX_FILES_PAGE_SIZE})`,
      },
      skipToken: {
        type: 'string',
        description: 'Continuation token from a previous response to get the next page of results.',
      },
    },
    required: ['conversationId'],
  },
};

const uploadFileToolDefinition: Tool = {
  name: 'teams_upload_file',
  description: 'Upload a local file to the user\'s OneDrive "Microsoft Teams Chat Files" folder via the Microsoft Graph API. Returns the uploaded file\'s metadata including itemId, fileName, SharePoint URLs, and a filesProperty string. To upload and send an attachment in one step, use teams_send_message with attachments: [{ filePath }]. The returned filesProperty is the low-level chatsvc metadata, not an attachments argument. Streams files in upload-session fragments, supporting 2 GiB and larger without a fixed local size limit. Increase the MCP client tool timeout for large uploads. The file path refers to the local filesystem of the machine running the MCP server.',
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: 'Absolute or relative path to the local file to upload (e.g., "/path/to/document.pdf").',
      },
    },
    required: ['filePath'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleGetSharedFiles(
  input: z.infer<typeof GetSharedFilesInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getSharedFiles(input.conversationId, {
    pageSize: input.pageSize,
    skipToken: input.skipToken,
  });

  return handleApiResult(result, (value) => ({
    conversationId: value.conversationId,
    returned: value.returned,
    files: value.files,
    ...(value.skipToken ? { skipToken: value.skipToken, hasMore: true } : { hasMore: false }),
  }));
}

async function handleUploadFile(
  input: z.infer<typeof UploadFileInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await uploadFile(input.filePath);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      itemId: result.value.itemId,
      fileName: result.value.fileName,
      fileType: result.value.fileType,
      fileSize: result.value.fileSize,
      baseUrl: result.value.baseUrl,
      objectUrl: result.value.objectUrl,
      listItemUniqueId: result.value.listItemUniqueId,
      filesProperty: result.value.filesProperty,
      note: 'File uploaded to OneDrive. To upload and send files in one step, use teams_send_message with attachments: [{ filePath }].',
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const getSharedFilesTool: RegisteredTool<typeof GetSharedFilesInputSchema> = {
  definition: getSharedFilesToolDefinition,
  schema: GetSharedFilesInputSchema,
  handler: handleGetSharedFiles,
};

export const DownloadFileInputSchema = z.object({
  url: z.string().url(),
  outputPath: z.string().min(1),
});

export const downloadFileTool: RegisteredTool<typeof DownloadFileInputSchema> = {
  definition: {
    name: 'teams_download_file',
    description: 'Download a Teams file using a SharePoint/OneDrive webUrl from teams_get_shared_files, a sharing or Doc.aspx viewer link, or a Microsoft Graph drive-item URL. Saves raw bytes to an absolute local outputPath and returns fileName, size, content type and SHA-256. Parent directory must exist; existing files are never overwritten. Streams to disk without a fixed size limit and cancels after 30 seconds without progress, removing incomplete files. Respects owner download blocks. Uses optional Microsoft Graph access from the Teams session; a Graph access failure does not trigger Teams login. For long downloads, increase the MCP client tool timeout.',
    inputSchema: {
      type: 'object',
      properties: {
        url: { type: 'string', description: 'The File item webUrl returned by teams_get_shared_files.' },
        outputPath: { type: 'string', description: 'Absolute destination file path on the MCP server machine; must not already exist.' },
      },
      required: ['url', 'outputPath'],
    },
  },
  schema: DownloadFileInputSchema,
  handler: async (input) => handleApiResult(await downloadFile(input.url, input.outputPath), value => ({ fileName: value.name, outputPath: value.outputPath, size: value.size, contentType: value.contentType, sha256: value.sha256 })),
};

export const uploadFileTool: RegisteredTool<typeof UploadFileInputSchema> = {
  definition: uploadFileToolDefinition,
  schema: UploadFileInputSchema,
  handler: handleUploadFile,
};

export const DownloadImageInputSchema = z.object({
  conversationId: z.string().min(1),
  messageId: z.string().min(1),
  imageIndex: z.number().int().min(0).default(0),
  outputPath: z.string().min(1),
});

export const downloadImageTool: RegisteredTool<typeof DownloadImageInputSchema> = {
  definition: {
    name: 'teams_download_image',
    description: 'Download an inline image from a Teams message. First read teams_get_message or teams_get_thread: their images array includes a zero-based index and downloadable flag. Re-fetches the message and downloads the selected Teams ASM image with existing Teams authentication, without Graph access. Saves original response bytes to an absolute local outputPath; parent directory must exist and existing files are never overwritten. Returns contentType, size and SHA-256. Redirects and unsupported image hosts are rejected; interrupted downloads are removed. The file extension should match the returned contentType, not the HTML image label.',
    inputSchema: {
      type: 'object',
      properties: {
        conversationId: { type: 'string', description: 'Conversation containing the image message.' },
        messageId: { type: 'string', description: 'Message id returned by teams_get_message or teams_get_thread.' },
        imageIndex: { type: 'integer', minimum: 0, default: 0, description: 'Zero-based index from the message images array (default 0).' },
        outputPath: { type: 'string', description: 'Absolute destination path on the MCP server machine; must not exist.' },
      },
      required: ['conversationId', 'messageId', 'outputPath'],
    },
  },
  schema: DownloadImageInputSchema,
  handler: async input => handleApiResult(await downloadImage(input.conversationId, input.messageId, input.imageIndex, input.outputPath), value => ({ ...value })),
};

/** All file-related tools. */
export const fileTools = [getSharedFilesTool, downloadFileTool, downloadImageTool, uploadFileTool];
