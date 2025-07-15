import { describe, it, expect, vi } from 'vitest';
import express from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { registerAuthTools } from '../src/auth-tools.js';
import { microsoftBearerTokenAuthMiddleware } from '../src/lib/microsoft-auth.js';

function createServer() {
  const authManager = {
    acquireTokenByDeviceCode: vi.fn().mockImplementation((cb: (t: string) => void) => {
      cb('login instructions');
      return Promise.resolve();
    }),
    testLogin: vi.fn().mockResolvedValue({ success: false, message: 'not logged' }),
    logout: vi.fn(),
  } as any;

  const server = new McpServer({ name: 'test', version: '1.0.0' });
  registerAuthTools(server, authManager);

  const app = express();
  app.use(express.json());

  const options = { enableAuthTools: true };
  const preAuth = (req: any, res: any, next: any) => {
    const toolName = req.body?.method === 'tools/call' ? req.body?.params?.name : undefined;
    if (options.enableAuthTools && toolName && ['login', 'logout', 'verify-login'].includes(toolName)) {
      return next();
    }
    return microsoftBearerTokenAuthMiddleware(req, res, next);
  };

  app.post('/mcp', preAuth, async (req, res) => {
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });
    res.on('close', () => transport.close());
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  });

  return app;
}

describe('HTTP auth tools', () => {
  it('allows login without Authorization when enabled', async () => {
    const app = createServer();
    const httpServer: any = await new Promise((resolve) => {
      const s = app.listen(0, () => resolve(s));
    });
    const port = httpServer.address().port;

    const payload = {
      jsonrpc: '2.0',
      id: 1,
      method: 'tools/call',
      params: { name: 'login', arguments: { force: false } },
    };

    const response = await new Promise<{ status: number; body: string }>((resolve, reject) => {
      const http = require('http');
      const req = http.request(
        {
          method: 'POST',
          hostname: 'localhost',
          port,
          path: '/mcp',
          headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json, text/event-stream',
          },
        },
        (res: any) => {
          let data = '';
          res.on('data', (chunk: any) => (data += chunk));
          res.on('end', () => resolve({ status: res.statusCode, body: data }));
        }
      );
      req.on('error', reject);
      req.write(JSON.stringify(payload));
      req.end();
    });

    expect(response.status).toBe(200);
    const parsed = JSON.parse(response.body);
    expect(parsed.result.content[0].text).toBe('login instructions');

    await new Promise((r) => httpServer.close(r));
  });
});
