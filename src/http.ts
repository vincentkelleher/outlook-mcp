import { createServer as createHttpServer } from "node:http";
import { createMcpHandler } from "@modelcontextprotocol/server";
import {
  localhostHostValidation,
  localhostOriginValidation,
  toNodeHandler,
} from "@modelcontextprotocol/node";
import { createServer } from "./mcp.ts";
import type { ToolDeps } from "./tools.ts";

/** Local-only Streamable HTTP; every request gets a fresh MCP server. */
export async function serveHttp(deps: ToolDeps, port = 3000) {
  if (!Number.isInteger(port) || port < 0 || port > 65535)
    throw new Error("port must be an integer from 0 to 65535.");

  const handler = createMcpHandler(() => createServer(deps), {
    responseMode: "sse",
  });
  const nodeHandler = toNodeHandler(handler);
  const validateHost = localhostHostValidation();
  const validateOrigin = localhostOriginValidation();
  const server = createHttpServer((req, res) => {
    if (!validateHost(req, res) || !validateOrigin(req, res)) return;
    if (req.url?.split("?")[0] !== "/mcp") {
      res.writeHead(404).end();
      return;
    }
    void nodeHandler(req, res).catch(() => {
      // Do not expose request data, mail content or stack traces.
      if (!res.headersSent) res.writeHead(500).end();
      else res.destroy();
    });
  });

  try {
    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      server.listen(port, "127.0.0.1", () => {
        server.off("error", reject);
        resolve();
      });
    });
  } catch (error) {
    await handler.close();
    throw error;
  }

  const address = server.address();
  if (!address || typeof address === "string")
    throw new Error("HTTP server has no TCP address.");

  let closing: Promise<void> | undefined;
  return {
    url: `http://127.0.0.1:${address.port}/mcp`,
    close(): Promise<void> {
      return (closing ??= (async () => {
        // Stop accepting connections before aborting active SSE exchanges.
        const stopped = new Promise<void>((resolve, reject) => {
          server.close((error) => (error ? reject(error) : resolve()));
        });
        try {
          await handler.close();
        } finally {
          server.closeAllConnections();
        }
        await stopped;
      })());
    },
  };
}
